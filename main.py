"""excel2md — 批量将 input 文件夹下的 Excel 文件转换为 Markdown。

用法:
    python main.py [--input input] [--output output] [--state state.json] [--debug]

规则:
    - 只处理 .xlsx / .xls 文件，其余文件一律跳过。
    - 输出保留输入相对目录结构；每个 Excel 对应输出一个文件夹，
      文件夹内按 sheet 拆分输出独立的 .md 文件。
    - 绝不删除输出文件夹中的既有内容。
    - 状态文件记录每个输入文件的 SHA-256（mmap 计算）：
      仅当 无记录 / hash 变化 / 任一 sheet 输出缺失 / --debug 时重新生成。
    - 分类算法（表格/文档）升级后，修改 _xlsx_classify.CLASSIFIER_VERSION
      即可让整份缓存失效、强制全部重新生成。
"""

import argparse
import hashlib
import json
import mmap
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# markitdown 是重导入链（magika/onnxruntime/pandas，约 0.6s），只在真正需要转换时
# 懒加载；纯 hash 跳过路径不导入它，「无变化」运行因此接近瞬时。
# 本常量需与 _xlsx_classify.CLASSIFIER_VERSION 保持同步；转换路径上会断言校验，
# 若过期会强制全部重转（安全方向）并给出告警。
_FAST_CLASSIFIER_VERSION = 3

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


def _get_converters() -> dict[str, object]:
    """懒加载 markitdown 转换器（重导入链，仅在实际转换时调用）。

    返回 {".xlsx": XlsxConverter, ".xls": XlsConverter}。
    """
    from markitdown.converters import XlsxConverter, XlsConverter
    from markitdown.converters._xlsx_classify import CLASSIFIER_VERSION

    if CLASSIFIER_VERSION != _FAST_CLASSIFIER_VERSION:
        print(
            f"⚠ _FAST_CLASSIFIER_VERSION({_FAST_CLASSIFIER_VERSION}) 与 markitdown "
            f"({CLASSIFIER_VERSION}) 不一致，本次按新版本重转",
            file=sys.stderr,
        )
    return {".xlsx": XlsxConverter(), ".xls": XlsConverter()}


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
    args = parser.parse_args(argv)

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

    if state.get("classifier_version") != _FAST_CLASSIFIER_VERSION:
        print(f"分类算法已更新 (v{_FAST_CLASSIFIER_VERSION})，本次强制全部重新生成")

    excel_files = discover_excel_files(input_root)

    converters = None  # 懒加载（_get_converters），只有真正要转换时才导入 markitdown

    discovered: set[str] = set()
    converted = skipped = errors = 0
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
            or rec is None
            or rec.get("hash") != digest
            or not isinstance(known_sheets, list)
            or not sheets_output_complete(excel_dir, known_sheets)
        )

        if regenerate:
            if converters is None:
                converters = _get_converters()
            try:
                with open(src, "rb") as fh:
                    sheet_data = converters[src.suffix.lower()].convert_sheets_with_assets(
                        fh, debug=args.debug
                    )
                if not sheet_data:
                    raise ValueError("工作簿没有任何 sheet")
                filenames = unique_sheet_filenames(list(sheet_data))
                excel_dir.mkdir(parents=True, exist_ok=True)
                n_images = 0
                for sheet_name, data in sheet_data.items():
                    (excel_dir / f"{filenames[sheet_name]}.md").write_text(
                        data["markdown"], encoding="utf-8"
                    )
                    # 图片写进 excel_dir/assets/<sheet>/，md 用相对路径引用。
                    for rel, img_bytes in data["images"].items():
                        img_path = excel_dir / rel
                        img_path.parent.mkdir(parents=True, exist_ok=True)
                        img_path.write_bytes(img_bytes)
                        n_images += 1
                files_rec[rel_key] = {
                    "hash": digest,
                    "converted_at": datetime.now(timezone.utc).isoformat(),
                    "sheets": list(sheet_data),
                }
                converted += 1
                suffix = f", {n_images} 张图" if n_images else ""
                print(
                    f"[{i}/{len(excel_files)}] ✔ {rel} -> {excel_dir} "
                    f"({len(sheet_data)} 个 sheet{suffix})"
                )
            except Exception as e:  # noqa: BLE001
                errors += 1
                print(f"[{i}/{len(excel_files)}] ✘ {rel} 转换失败: {e}", file=sys.stderr)
        else:
            skipped += 1
            print(f"[{i}/{len(excel_files)}] = {rel} 跳过（hash 一致且输出完整）")

    # 清理已从 input 删除文件的记录；output 中的孤儿 .md 保留（不删任何输出内容）。
    state["files"] = {k: v for k, v in files_rec.items() if k in discovered}
    state["classifier_version"] = _FAST_CLASSIFIER_VERSION
    save_state(state_path, state)

    if not excel_files:
        print(f"（{input_root} 下没有 Excel 文件）")
    print(f"\n{converted} 转换, {skipped} 跳过, {errors} 错误")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
