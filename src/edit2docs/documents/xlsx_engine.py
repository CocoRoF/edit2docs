"""Deterministic XLSX building blocks (no LLM).

A *sheet spec* is the interchange format the generator LLM emits:

    sheets:
      - name: "매출 요약"
        headers: ["분기", "매출(억원)", "YoY"]
        rows:
          - ["1분기", 120, "+12%"]
          - ["2분기", 135, "+9%"]
        widths: [10, 14, 10]        # optional, characters
        number_formats: {"B": "#,##0"}   # optional, column letter -> format

``xlsx_from_spec`` renders it (styled header row, frozen panes, auto
widths); ``xlsx_outline`` / ``apply_xlsx_edits`` give agents the same
inspect-then-patch loop the DOCX/PPTX sides have.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

__all__ = [
    "xlsx_from_spec",
    "xlsx_outline",
    "xlsx_to_markdown",
    "apply_xlsx_edits",
    "XlsxEdit",
    "XlsxEditResult",
]

_HEADER_FILL = PatternFill("solid", fgColor="1B64DA")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ---------------------------------------------------------------------------
# Spec -> XLSX
# ---------------------------------------------------------------------------


def xlsx_from_spec(spec: dict[str, Any]) -> bytes:
    """Render a sheet spec into a styled .xlsx package.

    Raises ``ValueError`` on a structurally invalid spec (bilingual
    message) — the generator retries on that signal.
    """
    sheets = spec.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        raise ValueError(
            "sheet spec must contain a non-empty `sheets` list. "
            "sheet spec에는 비어있지 않은 `sheets` 목록이 필요합니다."
        )

    wb = Workbook()
    wb.remove(wb.active)
    for index, sheet in enumerate(sheets, start=1):
        if not isinstance(sheet, dict):
            raise ValueError(f"sheets[{index - 1}] must be a mapping")
        title = str(sheet.get("name") or f"Sheet{index}")[:31]
        ws = wb.create_sheet(title=title)

        headers = [str(h) for h in (sheet.get("headers") or [])]
        rows = sheet.get("rows") or []
        if headers:
            ws.append(headers)
            for cell in ws[1]:
                cell.fill = _HEADER_FILL
                cell.font = _HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = _BORDER
            ws.freeze_panes = "A2"
        for row in rows:
            if not isinstance(row, (list, tuple)):
                raise ValueError(f"rows entries must be lists (sheet {title!r})")
            ws.append(list(row))

        # Body borders.
        first_body = 2 if headers else 1
        for row_cells in ws.iter_rows(min_row=first_body):
            for cell in row_cells:
                cell.border = _BORDER

        # Column widths: explicit, else content-driven (capped).
        widths = sheet.get("widths") or []
        column_count = max(len(headers), *(len(r) for r in rows), 1) if rows or headers else 1
        for col in range(1, column_count + 1):
            letter = get_column_letter(col)
            if col <= len(widths) and widths[col - 1]:
                ws.column_dimensions[letter].width = float(widths[col - 1])
            else:
                longest = 0
                for row_cells in ws.iter_rows(min_col=col, max_col=col):
                    value = row_cells[0].value
                    if value is not None:
                        # CJK chars are ~2 cells wide.
                        text = str(value)
                        width = sum(2 if ord(ch) > 0x2E80 else 1 for ch in text)
                        longest = max(longest, width)
                ws.column_dimensions[letter].width = min(max(longest + 2, 8), 48)

        for letter, fmt in (sheet.get("number_formats") or {}).items():
            start = 2 if headers else 1
            for (cell,) in ws.iter_rows(
                min_col=ws[f"{letter}1"].column,
                max_col=ws[f"{letter}1"].column,
                min_row=start,
            ):
                cell.number_format = str(fmt)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# XLSX -> outline / markdown
# ---------------------------------------------------------------------------


def _load(content: bytes):
    try:
        return load_workbook(io.BytesIO(content), data_only=False)
    except Exception as exc:
        raise ValueError(
            f"XLSX could not be opened: {exc}. XLSX 파일을 열 수 없습니다."
        ) from exc


def xlsx_outline(content: bytes, *, sample_rows: int = 8) -> dict:
    """Workbook structure: sheets, dimensions, headers, sample rows."""
    wb = _load(content)
    sheets = []
    for ws in wb.worksheets:
        rows = ws.max_row if ws.max_row else 0
        cols = ws.max_column if ws.max_column else 0
        sample = []
        for row_cells in ws.iter_rows(min_row=1, max_row=min(rows, sample_rows)):
            sample.append([cell.value for cell in row_cells])
        sheets.append(
            {
                "name": ws.title,
                "rows": rows,
                "columns": cols,
                "sample": sample,
            }
        )
    return {"sheets": sheets}


def xlsx_to_markdown(content: bytes, *, max_rows: int = 40) -> str:
    """Every sheet as a markdown table (row-capped for prompt budgets)."""
    wb = _load(content)
    parts: list[str] = []
    for ws in wb.worksheets:
        parts.append(f"## {ws.title}")
        rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, max_rows)))
        if not rows:
            parts.append("(empty)")
            continue
        table_lines = []
        for i, row_cells in enumerate(rows):
            cells = ["" if c.value is None else str(c.value) for c in row_cells]
            table_lines.append("| " + " | ".join(cells) + " |")
            if i == 0:
                table_lines.append("|" + "---|" * len(cells))
        if (ws.max_row or 0) > max_rows:
            table_lines.append(f"… ({ws.max_row - max_rows} more rows)")
        parts.append("\n".join(table_lines))
    return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# Targeted edits
# ---------------------------------------------------------------------------

_CELL_REF = re.compile(r"^[A-Za-z]{1,3}[1-9]\d*$")


@dataclass
class XlsxEdit:
    """One operation. ``action``:

    * ``set_cell`` — write ``value`` into ``cell`` (e.g. "B3") of ``sheet``.
      ``old_value`` (stringified compare) guards staleness.
    * ``append_rows`` — append ``rows`` (list of lists) to ``sheet``.
    * ``add_sheet`` — create sheet ``sheet`` with optional ``headers``/``rows``.
    """

    action: str
    sheet: str = ""
    cell: str | None = None
    value: Any = None
    old_value: Any = None
    rows: list | None = None
    headers: list | None = None


@dataclass
class XlsxEditResult:
    action: str
    status: str  # applied | stale | not_found | invalid
    message: str = ""


def apply_xlsx_edits(content: bytes, edits: Iterable[XlsxEdit]) -> tuple[bytes, list[XlsxEditResult]]:
    wb = _load(content)
    results: list[XlsxEditResult] = []
    for edit in edits:
        results.append(_apply_one(wb, edit))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), results


def _apply_one(wb, edit: XlsxEdit) -> XlsxEditResult:
    if edit.action == "add_sheet":
        if not edit.sheet:
            return XlsxEditResult(edit.action, "invalid", "add_sheet needs a sheet name")
        if edit.sheet in wb.sheetnames:
            return XlsxEditResult(edit.action, "invalid", f"sheet {edit.sheet!r} already exists")
        ws = wb.create_sheet(title=str(edit.sheet)[:31])
        if edit.headers:
            ws.append(list(edit.headers))
        for row in edit.rows or []:
            ws.append(list(row))
        return XlsxEditResult(edit.action, "applied")

    if edit.sheet not in wb.sheetnames:
        return XlsxEditResult(
            edit.action, "not_found",
            f"sheet {edit.sheet!r} not in {wb.sheetnames}",
        )
    ws = wb[edit.sheet]

    if edit.action == "set_cell":
        if not edit.cell or not _CELL_REF.match(edit.cell):
            return XlsxEditResult(edit.action, "invalid", f"bad cell ref {edit.cell!r}")
        target = ws[edit.cell.upper()]
        if edit.old_value is not None:
            current = "" if target.value is None else str(target.value)
            if current.strip() != str(edit.old_value).strip():
                return XlsxEditResult(edit.action, "stale", "cell value changed; refresh")
        target.value = edit.value
        return XlsxEditResult(edit.action, "applied")

    if edit.action == "append_rows":
        if not edit.rows:
            return XlsxEditResult(edit.action, "invalid", "append_rows needs rows")
        for row in edit.rows:
            if not isinstance(row, (list, tuple)):
                return XlsxEditResult(edit.action, "invalid", "rows entries must be lists")
            ws.append(list(row))
        return XlsxEditResult(edit.action, "applied")

    return XlsxEditResult(edit.action, "invalid", f"unknown action {edit.action!r}")
