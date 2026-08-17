"""excel2md — 批量将 input 文件夹下的 Excel 文件转换为 Markdown。

用法:
    python main.py [--input input] [--output output] [--state state.json] [--workers N] [--debug]

规则:
    - 只处理 .xlsx / .xls 文件，其余文件一律跳过。
    - 输出保留输入相对目录结构；每个 Excel 对应输出一个文件夹，
      文件夹内按 sheet 拆分输出独立的 .md 文件。
    - 绝不删除输出文件夹中的既有内容。
    - 状态文件记录每个输入文件的 SHA-256（mmap 计算）：
      仅当 无记录 / hash 变化 / 任一 sheet 输出缺失 / --debug 时重新生成。
    - 每个 (Excel, sheet) 都是独立任务，默认用 min(cpu_count, 8) 个进程并行转换；
      --workers 1 可切回纯串行。
    - 分类算法（表格/文档）升级后，修改 _xlsx_classify.CLASSIFIER_VERSION
      即可让整份缓存失效、强制全部重新生成。
"""

import argparse
import hashlib
import json
import mmap
import multiprocessing
import os
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

# markitdown 是重导入链（magika/onnxruntime/pandas，约 0.6s），只在真正需要转换时
# 懒加载；纯 hash 跳过路径不导入它，「无变化」运行因此接近瞬时。
# 本常量需与 _xlsx_classify.CLASSIFIER_VERSION 保持同步；转换路径上会断言校验，
# 若过期会强制全部重转（安全方向）并给出告警。
_FAST_CLASSIFIER_VERSION = 4

DEFAULT_INPUT = "input"
DEFAULT_OUTPUT = "output"
DEFAULT_STATE = "state.json"
EXCEL_SUFFIXES = {".xlsx", ".xls"}

_WINDOWS_RESERVED_NAMES = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
_INVALID_FILENAME_CHARS = set(r'<>:"/\|?*') | {chr(c) for c in range(32)}


def sanitize_filename(name: str) -> str:
    """sheet 名 → 安全的文件名（兼容 Windows）。"""
    out = "".join("_" if ch in _INVALID_FILENAME_CHARS else ch for ch in name)
    out = out.strip().rstrip(".").rstrip()
    if out.split(".", 1)[0].upper() in _WINDOWS_RESERVED_NAMES:
        out = "_" + out
    return out or "Sheet"


def unique_sheet_filenames(sheet_names: list[str]) -> dict[str, str]:
    """{原始 sheet 名: 文件名}，消除 sanitize 后的重名。"""
    used: set[str] = set()
    mapping: dict[str, str] = {}
    for name in sheet_names:
        base = sanitize_filename(name)
        filename = base
        i = 1
        while filename in used:
            i += 1
            filename = f"{base}_{i}"
        used.add(filename)
        mapping[name] = filename
    return mapping


def file_hash(path: Path) -> str:
    """用 mmap 计算文件 SHA-256，避免把整个文件读入 Python 内存。

    空文件单独处理——Windows 上 mmap 长度为 0 的文件会抛错。
    """
    with open(path, "rb") as fh:
        if fh.seek(0, os.SEEK_END) == 0:
            return hashlib.sha256(b"").hexdigest()
        fh.seek(0)
        with mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            return hashlib.sha256(mm).hexdigest()


def discover_excel_files(input_root: Path) -> list[Path]:
    """递归收集输入文件夹下的 Excel 文件（只按扩展名过滤，其余一概不碰）。"""
    return sorted(
        p
        for p in input_root.rglob("*")
        if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES
    )


def _discover_sheet_names(path: Path, suffix: str) -> list[str]:
    """有序 sheet 名列表，不做整簿解析。

    .xlsx 用 openpyxl 的 read_only load（11MB 文件实测 ~0.02s，完整 load ~4s）；
    .xls 用 xlrd 的 sheet_names()。懒导入，保持纯跳过路径不导入 markitdown。
    """
    if suffix == ".xlsx":
        import openpyxl

        with open(path, "rb") as fh:
            wb = openpyxl.load_workbook(fh, read_only=True, data_only=True)
            try:
                return list(wb.sheetnames)
            finally:
                wb.close()
    import xlrd

    with open(path, "rb") as fh:
        book = xlrd.open_workbook(file_contents=fh.read())
    return list(book.sheet_names())


def _worker_sheet_task(task: dict) -> dict:
    """在 worker 进程中转换一个 (file, sheet) 任务。

    ``task`` 只含可 pickling 的纯类型：{rel_key, sheet_name, filename,
    src, suffix, debug}。markitdown 在此懒导入（每进程一次，sys.modules 缓存）；
    父进程的纯跳过路径永不导入它。返回该 sheet 的
    ``{"markdown": str, "images": {relpath: bytes}}``，由父进程统一写盘。
    """
    from markitdown.converters import XlsxConverter, XlsConverter

    conv = XlsxConverter() if task["suffix"] == ".xlsx" else XlsConverter()
    with open(task["src"], "rb") as fh:
        return conv.convert_sheet_with_assets(fh, task["sheet_name"], debug=task["debug"])


def load_state(path: Path) -> dict:
    if not path.exists():
        return {"version": 1, "classifier_version": _FAST_CLASSIFIER_VERSION, "files": {}}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        # 直接报错退出，而不是静默丢弃用户的状态文件。
        raise SystemExit(f"状态文件损坏或无法读取: {path} ({e})") from e
    state.setdefault("version", 1)
    state.setdefault("classifier_version", _FAST_CLASSIFIER_VERSION)
    state.setdefault("files", {})
    return state


def save_state(path: Path, state: dict) -> None:
    """原子写：先写同目录 tmp 文件，再 os.replace（Windows 上可用）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def sheets_output_complete(excel_dir: Path, sheet_names: list[str]) -> bool:
    """该 Excel 的每个 sheet 输出 .md 是否都已存在。"""
    if not excel_dir.is_dir():
        return False
    filenames = unique_sheet_filenames(sheet_names)
    return all((excel_dir / f"{filenames[s]}.md").is_file() for s in sheet_names)


def main(argv: list[str] | None = None) -> int:
    # Windows 控制台默认编码（cp1252/cp936）无法编码 ✔/✘/⚠ 等字符，打印会抛
    # UnicodeEncodeError；强制 stdout/stderr 用 UTF-8（errors=replace 兜底）。
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(
        description="批量将 input 文件夹下的 Excel 转换为 Markdown 输出到 output",
    )
    parser.add_argument(
        "--input",
        default=DEFAULT_INPUT,
        help=f"输入文件夹（默认 {DEFAULT_INPUT}，不存在会自动创建）",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出文件夹（默认 {DEFAULT_OUTPUT}，不存在会自动创建）",
    )
    parser.add_argument(
        "--state",
        default=DEFAULT_STATE,
        help=f"状态文件路径（默认 {DEFAULT_STATE}）",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="总是重新生成，并打印每个 sheet 的表格/文档分类结果",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=min(os.cpu_count() or 1, 8),
        help="并行 worker 数（默认 min(cpu_count,8)；1 表示纯串行）",
    )
    args = parser.parse_args(argv)
    workers = max(1, args.workers)

    input_root = Path(args.input).resolve()
    output_root = Path(args.output).resolve()
    state_path = Path(args.state).resolve()

    # 自动创建默认的输入/输出文件夹。
    input_root.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    if output_root.is_relative_to(input_root) or input_root.is_relative_to(output_root):
        print(
            f"⚠ 提示: 输出文件夹 {output_root} 位于输入文件夹 {input_root} 内，"
            "扫描会忽略生成的 .md 文件，但仍建议分开",
            file=sys.stderr,
        )

    state = load_state(state_path)
    files_rec = state.setdefault("files", {})

    classifier_changed = state.get("classifier_version") != _FAST_CLASSIFIER_VERSION
    if classifier_changed:
        print(f"分类算法已更新 (v{_FAST_CLASSIFIER_VERSION})，本次强制全部重新生成")

    excel_files = discover_excel_files(input_root)

    # ---- Pass 1（串行、快）：hash → 跳过判定 → sheet 发现 → 建 (file, sheet) 任务 ----
    discovered: set[str] = set()
    skipped = errors = 0
    tasks: list[dict] = []
    rel_dir: dict[str, Path] = {}               # rel_key -> excel_dir
    rel_display: dict[str, Path] = {}           # rel_key -> rel（提示用）
    rel_digest: dict[str, str] = {}             # rel_key -> sha256
    rel_sheet_names: dict[str, list[str]] = {}  # rel_key -> 有序 sheet 列表
    rel_sheet_count: dict[str, int] = {}
    rel_image_count: dict[str, int] = {}

    for i, src in enumerate(excel_files, 1):
        rel = src.relative_to(input_root)
        rel_key = rel.as_posix()
        excel_dir = output_root / rel.with_suffix("")
        discovered.add(rel_key)

        try:
            digest = file_hash(src)
        except OSError as e:
            errors += 1
            print(f"[{i}/{len(excel_files)}] ✘ {rel} 无法读取: {e}", file=sys.stderr)
            continue

        rec = files_rec.get(rel_key)
        known_sheets = rec.get("sheets") if isinstance(rec, dict) else None
        regenerate = (
            args.debug
            or classifier_changed  # 分类版本升级后强制全量重转
            or rec is None
            or rec.get("hash") != digest
            or not isinstance(known_sheets, list)
            or not sheets_output_complete(excel_dir, known_sheets)
        )

        if not regenerate:
            skipped += 1
            print(f"[{i}/{len(excel_files)}] = {rel} 跳过（hash 一致且输出完整）")
            continue

        try:
            sheet_names = _discover_sheet_names(src, src.suffix.lower())
        except Exception as e:  # noqa: BLE001 — 损坏文件：记一次错误，跳过该文件继续
            errors += 1
            print(f"[{i}/{len(excel_files)}] ✘ {rel} 无法读取工作簿: {e}", file=sys.stderr)
            continue
        if not sheet_names:
            errors += 1
            print(f"[{i}/{len(excel_files)}] ✘ {rel} 工作簿没有任何 sheet", file=sys.stderr)
            continue
        # 文件名去重在父进程按整簿有序列表算一次——worker 只处理单 sheet，无法独立算对。
        filenames = unique_sheet_filenames(sheet_names)
        rel_dir[rel_key] = excel_dir
        rel_display[rel_key] = rel
        rel_digest[rel_key] = digest
        rel_sheet_names[rel_key] = sheet_names
        rel_sheet_count[rel_key] = len(sheet_names)
        rel_image_count[rel_key] = 0
        for sn in sheet_names:
            tasks.append(
                {
                    "rel_key": rel_key,
                    "sheet_name": sn,
                    "filename": filenames[sn],
                    "src": str(src),
                    "suffix": src.suffix.lower(),
                    "debug": args.debug,
                }
            )

    # 无任务（全部跳过 / 没有 Excel）：纯跳过路径不导入 markitdown，接近瞬时。
    if not tasks:
        state["files"] = {k: v for k, v in files_rec.items() if k in discovered}
        state["classifier_version"] = _FAST_CLASSIFIER_VERSION
        save_state(state_path, state)
        if not excel_files:
            print(f"（{input_root} 下没有 Excel 文件）")
        print(f"\n0 转换, {skipped} 跳过, {errors} 错误")
        return 1 if errors else 0

    # markitdown 侧版本校验（仅当确实要转换时才导入）。
    from markitdown.converters._xlsx_classify import CLASSIFIER_VERSION

    if CLASSIFIER_VERSION != _FAST_CLASSIFIER_VERSION:
        print(
            f"⚠ _FAST_CLASSIFIER_VERSION({_FAST_CLASSIFIER_VERSION}) 与 markitdown "
            f"({CLASSIFIER_VERSION}) 不一致，本次按新版本重转",
            file=sys.stderr,
        )

    # ---- Pass 2（并行转换 + 父进程统一写盘）----
    pending: dict[str, set[str]] = {}
    for t in tasks:
        pending.setdefault(t["rel_key"], set()).add(t["sheet_name"])
    failed_files: set[str] = set()
    converted = 0

    def mark_file_failed(rel_key: str, exc: Exception) -> None:
        nonlocal errors
        if rel_key in failed_files:
            return  # 一个文件只记一次错误
        failed_files.add(rel_key)
        errors += 1
        print(f"✘ {rel_display[rel_key]} 转换失败: {exc}", file=sys.stderr)

    def process_result(task: dict, result: dict) -> None:
        """写盘一个 sheet 的结果；某文件全部 sheet 成功后才 commit 状态记录。"""
        nonlocal converted
        rel_key = task["rel_key"]
        if rel_key in failed_files:
            return
        excel_dir = rel_dir[rel_key]
        try:
            excel_dir.mkdir(parents=True, exist_ok=True)
            (excel_dir / f"{task['filename']}.md").write_text(
                result["markdown"], encoding="utf-8"
            )
            # 图片写进 excel_dir/assets/<sheet>/，md 用相对路径引用。
            for img_rel, img_bytes in result["images"].items():
                img_path = excel_dir / img_rel
                img_path.parent.mkdir(parents=True, exist_ok=True)
                img_path.write_bytes(img_bytes)
                rel_image_count[rel_key] += 1
        except Exception as e:  # noqa: BLE001 — 写盘失败按整文件失败处理
            mark_file_failed(rel_key, e)
            return
        pending[rel_key].discard(task["sheet_name"])
        if not pending[rel_key]:
            # 全部 sheet 成功 → 此刻才提交状态记录（半提交状态永不被观察到）。
            files_rec[rel_key] = {
                "hash": rel_digest[rel_key],
                "converted_at": datetime.now(timezone.utc).isoformat(),
                "sheets": rel_sheet_names[rel_key],
            }
            converted += 1
            n_img = rel_image_count.get(rel_key, 0)
            suffix = f", {n_img} 张图" if n_img else ""
            print(
                f"✔ {rel_display[rel_key]} -> {excel_dir} "
                f"({rel_sheet_count[rel_key]} 个 sheet{suffix})"
            )

    if len(tasks) <= 1 or workers <= 1:
        for task in tasks:  # 单任务/串行：进程内直接跑，免起 pool 的开销
            try:
                process_result(task, _worker_sheet_task(task))
            except Exception as e:  # noqa: BLE001
                mark_file_failed(task["rel_key"], e)
    else:
        with ProcessPoolExecutor(max_workers=min(workers, len(tasks))) as pool:
            futures = {pool.submit(_worker_sheet_task, task): task for task in tasks}
            for fut in as_completed(futures):
                task = futures[fut]
                try:
                    process_result(task, fut.result())
                except Exception as e:  # noqa: BLE001
                    mark_file_failed(task["rel_key"], e)

    # 清理已从 input 删除文件的记录；output 中的孤儿 .md 保留（不删任何输出内容）。
    state["files"] = {k: v for k, v in files_rec.items() if k in discovered}
    state["classifier_version"] = _FAST_CLASSIFIER_VERSION
    save_state(state_path, state)

    if not excel_files:
        print(f"（{input_root} 下没有 Excel 文件）")
    print(f"\n{converted} 转换, {skipped} 跳过, {errors} 错误")
    return 1 if errors else 0


if __name__ == "__main__":
    multiprocessing.freeze_support()  # Nuitka/Windows spawn 必需
    sys.exit(main())
