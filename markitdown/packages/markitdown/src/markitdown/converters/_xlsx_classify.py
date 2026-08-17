"""Sheet classification and region-based rendering helpers for the xlsx/xls converters.

Two tiers of rendering:

- **Region-based mixed rendering** (``detect_regions`` + ``render_region_mixed``):
  within one sheet, bordered/dense regions render as Markdown tables and the
  surrounding prose renders as document paragraphs. Planners draw borders around
  data tables and leave prose borderless, so the border signal is primary; a
  fully-nonempty 4×4 window catches borderless dense tables.
- **Whole-sheet fallback** (``classify_sheet`` + ``render_document_sheet``): used
  when a sheet has no table regions (pure document, or ``.xls`` which has no
  border information).

All thresholds are intentionally tunable — bump ``CLASSIFIER_VERSION`` after
changing them so the batch tool's conversion cache invalidates.
"""

from __future__ import annotations

from typing import Any, NamedTuple

# Bump whenever the classification heuristic or renderer changes: the excel2md
# batch tool treats a mismatched version as "reconvert everything".
CLASSIFIER_VERSION = 4

# --- Whole-sheet classification thresholds ------------------------------------
DOC_MAX_COLS = 2                 # ≤2 columns is a strong document signal
DOC_MAX_COLS_WITH_KEYWORD = 4    # ≤4 columns + a doc-ish sheet name → document
MERGED_RATIO_THRESHOLD = 0.15    # fraction of all cells that are merged
DOC_AVG_TEXT_LEN = 12            # avg non-empty cell text length (sparse sheets)
DOC_LONG_TEXT_RATIO = 0.03       # fraction of text cells longer than 50 chars
DOC_SPARSE_NONEMPTY_RATIO = 0.5  # non-empty ratio below this + long text → document
DOC_LABEL_MAX_LEN = 40           # first cell treated as a label if shorter than this
DOC_KEYWORDS = (
    "策划",
    "方案",
    "说明",
    "需求",
    "背景",
    "概述",
    "综述",
    "简介",
    "设计",
    "文档",
    "note",
    "readme",
    "description",
)

# --- Region-based table detection thresholds -----------------------------------
TABLE_BORDER_RUN = 3             # bordered cells in a row → table row (xlsx only)
TABLE_CONTENT_RUN = 5            # contiguous non-empty cells → table row
DENSE_WINDOW_ROWS = 4            # fully-nonempty window → all rows in it are table rows
DENSE_WINDOW_COLS = 4
CAPTION_RATIO = 0.6              # header = first row with in-span count >= ratio * region_max
MIN_TABLE_WIDTH = 3              # region renders as a table only if max in-span width >= this
GAP_TOLERANCE = 0                # one non-table row breaks a region
CAPTION_BOLD = True              # single-cell captions render as **bold**
EMIT_OUT_OF_SPAN = True          # emit long out-of-span cells in table rows as a blockquote
OUT_OF_SPAN_MIN_LEN = 10

MODE_TABLE = "table"
MODE_DOCUMENT = "document"
MODE_MIXED = "mixed"

# --- Filename helpers ----------------------------------------------------------
# 与 main.py 中的同名实现保持一致（main.py 为保持跳过路径不导入 markitdown，
# 在本地保留了一份副本，两处改动需同步）。
_WINDOWS_RESERVED_NAMES = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
_INVALID_FILENAME_CHARS = set(r'<>:"/\|?*') | {chr(c) for c in range(32)}


def sanitize_filename(name: str) -> str:
    """sheet/asset 名 → 安全的文件名（兼容 Windows）。"""
    out = "".join("_" if ch in _INVALID_FILENAME_CHARS else ch for ch in name)
    out = out.strip().rstrip(".").rstrip()
    if out.split(".", 1)[0].upper() in _WINDOWS_RESERVED_NAMES:
        out = "_" + out
    return out or "Sheet"


def _cell_text(value: Any) -> str:
    """Normalize a raw cell value to stripped text ('' means empty).

    Python floats (openpyxl formula-cached values like ``8.440000000000001``)
    are formatted with ``%.15g`` to match Excel General's ~15 significant
    digits, so the floating-point noise never leaks into the markdown. bool is
    a subclass of int (not float); the explicit guard keeps ``True``/``False``
    as text. A pandas-style ``"nan"`` string still renders empty.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        text = str(value).strip()
    elif isinstance(value, float):
        text = f"{value:.15g}"
    else:
        text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


class RowMeta(NamedTuple):
    """Per-row features used to decide whether the row belongs to a table region."""

    idx: int
    content_count: int
    max_content_run: int
    border_run: int
    border_lo: int | None  # 0-based first bordered column
    border_hi: int | None  # 0-based last bordered column


def analyze_rows(cells: list[list[Any]], border_flags: list[list[bool]] | None = None) -> list[RowMeta]:
    """One pass over the sheet producing a :class:`RowMeta` per row.

    ``border_flags`` is ``None`` for ``.xls`` (no border info); every row then
    has zero borders.
    """
    metas: list[RowMeta] = []
    for r, row in enumerate(cells):
        content_count = 0
        max_content_run = 0
        border_run = 0
        cur = curb = 0
        border_lo = border_hi = None
        borders = border_flags[r] if border_flags else None
        for c, value in enumerate(row):
            if _cell_text(value):
                content_count += 1
                cur += 1
                max_content_run = max(max_content_run, cur)
            else:
                cur = 0
            if borders is not None and borders[c]:
                border_run += 1
                curb += 1
                if border_lo is None:
                    border_lo = c
                border_hi = c
            else:
                curb = 0
        metas.append(
            RowMeta(r, content_count, max_content_run, border_run, border_lo, border_hi)
        )
    return metas


def is_table_row(meta: RowMeta, dense_rows: set[int]) -> bool:
    """A row belongs to a table region if any signal fires.

    The border run counts bordered cells regardless of content (caption rows
    like 综述 r44 have 11 bordered cells but one content cell).
    """
    return (
        meta.border_run >= TABLE_BORDER_RUN
        or meta.max_content_run >= TABLE_CONTENT_RUN
        or meta.idx in dense_rows
    )


def dense_4x4_rows(
    cells: list[list[Any]],
    rows: int = DENSE_WINDOW_ROWS,
    cols: int = DENSE_WINDOW_COLS,
) -> set[int]:
    """Rows covered by a fully-nonempty ``rows x cols`` window.

    O(n·m) time, O(m) memory via per-column consecutive-run counters — cheap
    even for the 53k-row 熔金存量 sheet. Required for borderless dense tables
    (e.g. markitdown's ``test.xlsx``, a 4-column table with content runs of 4).
    """
    n = len(cells)
    m = len(cells[0]) if cells else 0
    if n < rows or m < cols:
        return set()
    dense: set[int] = set()
    col_stack = [0] * m  # consecutive non-empty cells per column up to current row
    for r in range(n):
        row = cells[r]
        for c in range(m):
            col_stack[c] = col_stack[c] + 1 if _cell_text(row[c]) else 0
        run = 0
        for c in range(m):
            run = run + 1 if col_stack[c] >= rows else 0
            if run >= cols:
                dense.update(range(r - rows + 1, r + 1))
                break
    return dense


def group_table_regions(table_flags: list[bool]) -> list[tuple[int, int]]:
    """Contiguous runs of table rows → ``(start, end)`` 0-based inclusive.

    Gap tolerance 0: a single non-table row breaks a region. This keeps
    separate tables apart even when they sit one borderless row apart
    (综述 rows 57–61 / 63–67; 数值 r34, r94, r136).
    """
    regions: list[tuple[int, int]] = []
    i = 0
    n = len(table_flags)
    while i < n:
        if not table_flags[i]:
            i += 1
            continue
        j = i
        while j + 1 < n and table_flags[j + 1]:
            j += 1
        regions.append((i, j))
        i = j + 1
    return regions


def _content_extent(row: list[Any]) -> tuple[int, int] | None:
    cols = [c for c, v in enumerate(row) if _cell_text(v)]
    return (cols[0], cols[-1]) if cols else None


def _row_content_in_span(row: list[Any], lo: int, hi: int) -> int:
    return sum(1 for v in row[lo : hi + 1] if _cell_text(v))


def detect_region_layout(
    cells: list[list[Any]],
    metas: list[RowMeta],
    region: tuple[int, int],
    use_border: bool = True,
) -> tuple[int, int, int, list[int]] | None:
    """Return ``(span_lo, span_hi, header_idx, caption_idxs)`` for a table region.

    Two-pass refinement: the initial span is the union of border extent over all
    region rows; if the header is not the first row, the span is recomputed from
    the header..end rows only, so a wider caption border (综述 r44 borders C–M
    but its data content is C) does not widen the table.
    """
    r0, r1 = region
    if use_border:
        los = [m.border_lo for m in metas[r0 : r1 + 1] if m.border_lo is not None]
        his = [m.border_hi for m in metas[r0 : r1 + 1] if m.border_hi is not None]
    else:
        los, his = [], []
    if los:
        lo, hi = min(los), max(his)
    else:  # fallback: content extent (borderless tables / .xls)
        spans = [s for r in range(r0, r1 + 1) if (s := _content_extent(cells[r]))]
        if not spans:
            return None
        lo = min(s[0] for s in spans)
        hi = max(s[1] for s in spans)

    region_max = max(_row_content_in_span(cells[r], lo, hi) for r in range(r0, r1 + 1))
    header_idx = next(
        r
        for r in range(r0, r1 + 1)
        if _row_content_in_span(cells[r], lo, hi) >= CAPTION_RATIO * region_max
    )

    if use_border and header_idx > r0:
        los = [m.border_lo for m in metas[header_idx : r1 + 1] if m.border_lo is not None]
        his = [m.border_hi for m in metas[header_idx : r1 + 1] if m.border_hi is not None]
        if los:
            lo, hi = min(los), max(his)
            region_max = max(
                _row_content_in_span(cells[r], lo, hi) for r in range(r0, r1 + 1)
            )
            header_idx = next(
                r
                for r in range(r0, r1 + 1)
                if _row_content_in_span(cells[r], lo, hi) >= CAPTION_RATIO * region_max
            )

    captions = [r for r in range(r0, header_idx)]
    return lo, hi, header_idx, captions


def render_document_row(items: list[str]) -> str:
    """Render one row's cleaned cell texts as document prose."""
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2 and len(items[0]) <= DOC_LABEL_MAX_LEN:
        return f"**{items[0]}**: {items[1]}"
    return f"**{items[0]}**: {' '.join(items[1:])}"


def render_document_sheet(
    cells: list[list[Any]],
    images_by_row: dict[int, list[str]] | None = None,
) -> str:
    """Render a whole sheet as document prose (no table regions).

    ``images_by_row`` maps 0-based row index → list of Markdown image lines to
    emit right after that row (empty rows with only an image still emit it).
    """
    out: list[str] = []
    for r, row in enumerate(cells):
        items = [t for t in (_cell_text(v) for v in row) if t]
        if items:
            out.append(render_document_row(items))
            out.append("")
        if images_by_row:
            for line in images_by_row.get(r, []):
                out.append(line)
                out.append("")
    return "\n".join(out).strip()


def _render_caption(row: list[Any], lo: int, hi: int) -> str:
    items = [_cell_text(v) for v in row[lo : hi + 1] if _cell_text(v)]
    if not items:
        return ""
    if len(items) == 1 and CAPTION_BOLD:
        return f"**{items[0]}**"
    return render_document_row(items)


def _render_out_of_span(cells: list[list[Any]], a: int, b: int, lo: int, hi: int) -> str:
    """Long non-empty cells outside the table span → blockquote (prose notes)."""
    if not EMIT_OUT_OF_SPAN:
        return ""
    notes: list[str] = []
    for r in range(a, b + 1):
        for c, v in enumerate(cells[r]):
            text = _cell_text(v)
            if (c < lo or c > hi) and len(text) >= OUT_OF_SPAN_MIN_LEN:
                notes.append(f"> {text}")
    return "\n".join(notes)


def compute_sheet_stats(
    *,
    n_rows: int,
    n_cols: int,
    merged_count: int,
    sheet_name: str,
    cells: list[list[Any]],
) -> dict[str, Any]:
    """Compute whole-sheet classification features from raw sheet data."""
    total = max(1, n_rows * n_cols)
    non_empty = 0
    lengths: list[int] = []
    long_count = 0
    max_non_empty_in_row = 0
    for row in cells:
        non_empty_in_row = 0
        for value in row:
            text = _cell_text(value)
            if not text:
                continue
            non_empty += 1
            non_empty_in_row += 1
            lengths.append(len(text))
            if len(text) > 50:
                long_count += 1
        max_non_empty_in_row = max(max_non_empty_in_row, non_empty_in_row)

    n_text = len(lengths)
    return {
        "n_rows": n_rows,
        "n_cols": n_cols,
        "merged_ratio": (merged_count / total) if merged_count else 0.0,
        "non_empty_ratio": non_empty / total,
        "avg_text_len": (sum(lengths) / n_text) if n_text else 0.0,
        "max_text_len": max(lengths) if lengths else 0,
        "long_text_ratio": (long_count / n_text) if n_text else 0.0,
        "max_non_empty_in_row": max_non_empty_in_row,
        "sheet_name_has_keyword": any(
            keyword.lower() in str(sheet_name).lower() for keyword in DOC_KEYWORDS
        ),
    }


def classify_sheet(stats: dict[str, Any]) -> tuple[str, str]:
    """Return ``(mode, reason)``; mode ∈ {"table", "document"}.

    Whole-sheet fallback used when a sheet has no detected table regions.
    """
    # 1. Narrow sheets are documents (label:value / prose layout).
    if stats["n_cols"] <= DOC_MAX_COLS:
        return MODE_DOCUMENT, f"cols<={DOC_MAX_COLS} ({stats['n_cols']})"

    # 2. Heavily merged sheets (merged labels spanning cells) are documents.
    if stats["merged_ratio"] >= MERGED_RATIO_THRESHOLD:
        return MODE_DOCUMENT, f"merged_ratio={stats['merged_ratio']:.2f}"

    # 3. Doc-keyword sheet name + narrow-ish + never a dense wide row → document.
    if (
        stats["sheet_name_has_keyword"]
        and stats["n_cols"] <= DOC_MAX_COLS_WITH_KEYWORD
        and stats["max_non_empty_in_row"] <= 2
    ):
        return MODE_DOCUMENT, "sheet-keyword + narrow"

    # 4. Sparse, long-text sheets are documents (paragraph-style planning sheets).
    if (
        stats["non_empty_ratio"] <= DOC_SPARSE_NONEMPTY_RATIO
        and (
            stats["avg_text_len"] >= DOC_AVG_TEXT_LEN
            or stats["long_text_ratio"] >= DOC_LONG_TEXT_RATIO
        )
    ):
        return MODE_DOCUMENT, (
            f"long-text avg={stats['avg_text_len']:.1f} "
            f"sparse={stats['non_empty_ratio']:.2f}"
        )

    return MODE_TABLE, "default"
