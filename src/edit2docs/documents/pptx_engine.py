"""Structural PPTX editing — addressable slide model + surgical operations.

This is the PPTX counterpart of :mod:`docx_engine` / :mod:`xlsx_engine`: the
agent sees a compact **addressable outline** of each slide (shapes by
``cNvPr`` id, tables by cell, charts by series) and emits **precise operations**
that are applied **byte-preservingly** through contextifier's raw layer — only
the slide parts an edit actually touches are rewritten; headers, theme styling,
native charts/tables and everything untouched stay byte-identical.

This is deliberately the opposite of the legacy full-slide SVG-rewrite path in
:mod:`edit2docs.tools.edit_deck`, which renders a whole slide to flat SVG, asks
the model to redraw it, and converts back — flattening tables/charts and
dropping any element the model forgets. Here nothing is redrawn: a table cell,
a chart's data, or a paragraph's text is mutated in place.

Addressing (all indices as shown in :func:`pptx_outline`):

* ``slide``  — 1-based slide number.
* ``shape``  — the shape's ``cNvPr/@id`` (stable across a turn); also stamped
  into the preview SVG as ``data-e2p-shape`` / ``data-e2p-table``.
* ``para``   — paragraph index inside a text shape (0-based).
* ``row`` / ``col`` — table cell address.
* ``chart``  — chart ordinal within the slide (0-based).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass

__all__ = [
    "PptxEdit",
    "PptxEditResult",
    "pptx_outline",
    "apply_pptx_edits",
    "VALID_PPTX_ACTIONS",
]

#: Every action the surgical engine understands. The planner is constrained to
#: this set; anything else is reported ``invalid`` rather than applied.
VALID_PPTX_ACTIONS = (
    "set_text",          # replace paragraph `para` of shape `shape`
    "set_shape_style",   # restyle shape text (color/size/bold/italic) + fill
    "set_shape_position",  # move/resize shape `shape` (inches)
    "set_table_cell",    # replace cell (row,col) of table shape `shape`
    "set_cell_style",    # restyle cell (row,col): fill / font
    "insert_row",        # insert a row into table shape `shape` at `at`
    "delete_row",        # delete row `row` of table shape `shape`
    "set_chart_data",    # rewrite chart `chart`'s categories + series
    "set_chart_title",   # rewrite chart `chart`'s title
)

#: EMU per inch — the outline reports geometry in inches; ops accept inches.
_EMU_PER_INCH = 914400


# ---------------------------------------------------------------------------
# Addressable outline (read)
# ---------------------------------------------------------------------------


def pptx_outline(content: bytes) -> list[dict]:
    """Addressable structure of every slide, in slide + document order.

    Returns a flat list of entries the planner's ops resolve against:

    * text paragraph — ``{"slide", "shape", "para", "kind":"text", "name", "text"}``
      (one entry per non-empty paragraph of a text shape)
    * table header    — ``{"slide", "shape", "kind":"table", "name", "rows", "cols"}``
    * table cell      — ``{"slide", "table", "row", "col", "text"}``
      (``table`` is the table shape's id; only non-empty cells are listed)
    * chart           — ``{"slide", "chart", "kind":"chart", "title",
      "chart_type", "series":[...]}``
    * other shape     — ``{"slide", "shape", "kind", "name"}``
      (pictures / groups / diagrams / unknown — so the planner knows they exist)

    Best-effort and never raises: a malformed slide degrades to fewer entries
    rather than failing the whole outline.
    """
    from contextifier import open_raw

    raw = open_raw(content, extension="pptx")
    outline: list[dict] = []

    def _geometry_inches(info) -> dict | None:
        vals = (
            getattr(info, "left", None), getattr(info, "top", None),
            getattr(info, "width", None), getattr(info, "height", None),
        )
        if any(v is None for v in vals):
            return None
        return {
            "left": round(vals[0] / _EMU_PER_INCH, 2),
            "top": round(vals[1] / _EMU_PER_INCH, 2),
            "width": round(vals[2] / _EMU_PER_INCH, 2),
            "height": round(vals[3] / _EMU_PER_INCH, 2),
        }

    for s_idx, slide in enumerate(raw.slides):
        n = s_idx + 1
        # tables addressed by shape id — collect ids so text-shape listing can
        # skip the table graphicFrames (their text lives in cells).
        try:
            tables = {t.shape_id: t for t in slide.tables}
        except Exception:
            tables = {}
        try:
            shapes = slide.shapes
        except Exception:
            shapes = []
        for info in shapes:
            if info.id in tables:
                continue  # handled in the table pass below
            pos = _geometry_inches(info)
            if info.kind == "text":
                paras = (info.text or "").split("\n")
                emitted = False
                for p_idx, para_text in enumerate(paras):
                    if para_text.strip():
                        entry = {
                            "slide": n, "shape": info.id, "para": p_idx,
                            "kind": "text", "name": info.name, "text": para_text,
                        }
                        if p_idx == 0 and pos:
                            entry["pos"] = pos
                        outline.append(entry)
                        emitted = True
                if not emitted:
                    # an empty text placeholder is still addressable (para 0)
                    entry = {
                        "slide": n, "shape": info.id, "para": 0,
                        "kind": "text", "name": info.name, "text": "",
                    }
                    if pos:
                        entry["pos"] = pos
                    outline.append(entry)
            elif info.kind not in ("chart",):  # charts listed from the part below
                entry = {
                    "slide": n, "shape": info.id, "kind": info.kind,
                    "name": info.name,
                }
                if pos:
                    entry["pos"] = pos
                outline.append(entry)
        # tables: header + non-empty cells
        for shape_id, table in tables.items():
            try:
                rows, cols = table.n_rows, table.n_cols
            except Exception:
                continue
            outline.append({
                "slide": n, "shape": shape_id, "kind": "table",
                "name": "", "rows": rows, "cols": cols,
            })
            for r in range(rows):
                for c in range(cols):
                    try:
                        text = table.cell(r, c).text.strip()
                    except Exception:
                        continue
                    if text:
                        outline.append({
                            "slide": n, "table": shape_id,
                            "row": r, "col": c, "text": text,
                        })
        # charts (via contextifier ChartModel; read-only summary here)
        try:
            charts = slide.charts
        except Exception:
            charts = []
        for c_idx, chart in enumerate(charts):
            try:
                series = [
                    {"name": s.name, "categories": list(s.categories),
                     "values": list(s.values)}
                    for s in chart.series
                ]
            except Exception:
                series = []
            try:
                title, ctype = chart.title, chart.kind
            except Exception:
                title, ctype = None, None
            outline.append({
                "slide": n, "chart": c_idx, "kind": "chart",
                "title": title, "chart_type": ctype, "series": series,
            })
    return outline


# ---------------------------------------------------------------------------
# Operations
# ---------------------------------------------------------------------------


@dataclass
class PptxEdit:
    """One surgical operation on a slide. ``action`` ∈ :data:`VALID_PPTX_ACTIONS`.

    * ``set_text``       — ``slide, shape, para, new_text`` (``old_text`` guards)
    * ``set_table_cell`` — ``slide, shape, row, col, new_text`` (``old_text`` guards)
    * ``insert_row``     — ``slide, shape, at`` (row cloned as style template)
    * ``delete_row``     — ``slide, shape, row``
    * ``set_chart_data`` — ``slide, chart, categories, series``
    * ``set_chart_title``— ``slide, chart, title``
    """

    action: str
    slide: int                       # 1-based slide number
    shape: int | None = None         # cNvPr id (text / table ops)
    para: int = 0                    # paragraph index (set_text)
    row: int | None = None           # table cell / delete_row
    col: int | None = None           # table cell
    at: int | None = None            # insert_row position
    chart: int | None = None         # chart ordinal within slide
    new_text: str = ""
    old_text: str | None = None      # optimistic-concurrency guard
    title: str | None = None
    categories: list | None = None
    series: list | None = None       # [{"name": str, "values": [num, ...]}, ...]
    # set_shape_style
    color: str | None = None         # font color "RRGGBB"
    size_pt: float | None = None     # font size in points
    bold: bool | None = None
    italic: bool | None = None
    fill: str | None = None          # shape fill "RRGGBB"
    # set_shape_position — inches
    left: float | None = None
    top: float | None = None
    width: float | None = None
    height: float | None = None


@dataclass
class PptxEditResult:
    action: str
    status: str  # applied | stale | not_found | invalid | unsupported
    message: str = ""


def _normalize(text: str) -> str:
    return " ".join((text or "").split())


def _row_sort_key(edit: PptxEdit) -> tuple:
    """Row-shifting ops (insert/delete) on the same table must run highest-row
    first so earlier addresses stay valid — mirrors docx's descending sort."""
    idx = edit.at if edit.at is not None else (edit.row if edit.row is not None else -1)
    return (edit.slide, edit.shape if edit.shape is not None else -1, idx)


def apply_pptx_edits(
    content: bytes, edits: Iterable[PptxEdit]
) -> tuple[bytes, list[PptxEditResult]]:
    """Apply surgical edits losslessly via contextifier's raw layer.

    Only the slide parts an edit touches are rewritten; every other part of the
    package stays byte-identical. Per-edit soft failures never abort the batch —
    each op yields an ``applied | stale | not_found | invalid | unsupported``
    result, reported in caller order.
    """
    from contextifier import open_raw

    try:
        raw = open_raw(content, extension="pptx")
    except Exception as exc:
        raise ValueError(
            f"PPTX could not be opened: {exc}. PPTX 파일을 열 수 없습니다."
        ) from exc

    edit_list = list(edits)
    results: list[PptxEditResult | None] = [None] * len(edit_list)

    # Row insert/delete shift row indices within one table; run them
    # highest-row-first so each op's original address stays valid. Non-row ops
    # are order-independent (shape id / chart ordinal don't shift).
    ordered = sorted(enumerate(edit_list), key=lambda pair: _row_sort_key(pair[1]), reverse=True)
    for index, edit in ordered:
        results[index] = _apply_one(raw, edit)

    return raw.to_bytes(), [r for r in results if r is not None]


def _slide_of(raw, edit: PptxEdit):
    """Resolve the 1-based ``edit.slide`` to a RawSlide, or None if invalid."""
    slides = raw.slides
    if edit.slide is None or not (1 <= edit.slide <= len(slides)):
        return None
    return slides[edit.slide - 1]


def _apply_one(raw, edit: PptxEdit) -> PptxEditResult:
    if edit.action not in VALID_PPTX_ACTIONS:
        return PptxEditResult(edit.action, "invalid", f"unknown action {edit.action!r}")

    slide = _slide_of(raw, edit)
    if slide is None:
        return PptxEditResult(edit.action, "not_found", "slide index out of range")

    try:
        if edit.action == "set_text":
            return _apply_set_text(slide, edit)
        if edit.action == "set_shape_style":
            return _apply_set_shape_style(slide, edit)
        if edit.action == "set_shape_position":
            return _apply_set_shape_position(slide, edit)
        if edit.action == "set_table_cell":
            return _apply_set_table_cell(slide, edit)
        if edit.action == "set_cell_style":
            return _apply_set_cell_style(slide, edit)
        if edit.action in ("insert_row", "delete_row"):
            return _apply_row_op(slide, edit)
        if edit.action in ("set_chart_data", "set_chart_title"):
            return _apply_chart_op(slide, edit)
    except Exception as exc:  # last-resort: one bad op never kills the batch
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "invalid", f"unhandled action {edit.action!r}")


def _apply_set_text(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_text needs a shape id")
    try:
        current_body = slide.get_text(edit.shape)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except ValueError:
        return PptxEditResult(edit.action, "invalid", "shape has no editable text")
    if edit.old_text is not None:
        paras = current_body.split("\n")
        current = paras[edit.para] if 0 <= edit.para < len(paras) else ""
        if _normalize(current) != _normalize(edit.old_text):
            return PptxEditResult(edit.action, "stale", "paragraph text changed; refresh")
    try:
        slide.set_text(edit.shape, edit.new_text, para=edit.para)
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "paragraph index out of range")
    return PptxEditResult(edit.action, "applied")


def _apply_set_shape_style(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_shape_style needs a shape id")
    font_changed = any(
        v is not None for v in (edit.color, edit.size_pt, edit.bold, edit.italic)
    )
    if not font_changed and edit.fill is None:
        return PptxEditResult(edit.action, "invalid", "set_shape_style needs a style attribute")
    try:
        if font_changed:
            slide.set_shape_font(
                edit.shape,
                color=edit.color,
                size_pt=edit.size_pt,
                bold=edit.bold,
                italic=edit.italic,
                para=(edit.para if edit.para else None),
            )
        if edit.fill is not None:
            slide.set_shape_fill(edit.shape, edit.fill)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except ValueError as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")


def _apply_set_shape_position(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_shape_position needs a shape id")
    if all(v is None for v in (edit.left, edit.top, edit.width, edit.height)):
        return PptxEditResult(edit.action, "invalid", "set_shape_position needs a coordinate")

    def _emu(v):
        return int(round(v * _EMU_PER_INCH)) if v is not None else None

    try:
        slide.set_shape_position(
            edit.shape,
            left=_emu(edit.left), top=_emu(edit.top),
            width=_emu(edit.width), height=_emu(edit.height),
        )
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    return PptxEditResult(edit.action, "applied")


def _table_by_shape(slide, shape_id: int | None):
    if shape_id is None:
        return None
    for table in slide.tables:
        if table.shape_id == shape_id:
            return table
    return None


def _apply_set_table_cell(slide, edit: PptxEdit) -> PptxEditResult:
    table = _table_by_shape(slide, edit.shape)
    if table is None:
        return PptxEditResult(edit.action, "not_found", "no table with that shape id")
    if edit.row is None or edit.col is None:
        return PptxEditResult(edit.action, "invalid", "set_table_cell needs row+col")
    try:
        cell = table.cell(edit.row, edit.col)
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "table cell out of range")
    if edit.old_text is not None and _normalize(cell.text) != _normalize(edit.old_text):
        return PptxEditResult(edit.action, "stale", "cell text changed; refresh")
    cell.set_text(edit.new_text)
    return PptxEditResult(edit.action, "applied")


def _apply_set_cell_style(slide, edit: PptxEdit) -> PptxEditResult:
    table = _table_by_shape(slide, edit.shape)
    if table is None:
        return PptxEditResult(edit.action, "not_found", "no table with that shape id")
    if edit.row is None or edit.col is None:
        return PptxEditResult(edit.action, "invalid", "set_cell_style needs row+col")
    if all(v is None for v in (edit.fill, edit.color, edit.size_pt, edit.bold, edit.italic)):
        return PptxEditResult(edit.action, "invalid", "set_cell_style needs a style attribute")
    try:
        cell = table.cell(edit.row, edit.col)
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "table cell out of range")
    cell.set_style(
        fill=edit.fill, color=edit.color, size_pt=edit.size_pt,
        bold=edit.bold, italic=edit.italic,
    )
    return PptxEditResult(edit.action, "applied")


def _apply_row_op(slide, edit: PptxEdit) -> PptxEditResult:
    table = _table_by_shape(slide, edit.shape)
    if table is None:
        return PptxEditResult(edit.action, "not_found", "no table with that shape id")
    try:
        if edit.action == "insert_row":
            at = edit.at if edit.at is not None else table.n_rows
            table.insert_row(at)
        else:  # delete_row
            if edit.row is None:
                return PptxEditResult(edit.action, "invalid", "delete_row needs row")
            table.delete_row(edit.row)
    except (IndexError, ValueError) as exc:
        return PptxEditResult(edit.action, "not_found", str(exc))
    return PptxEditResult(edit.action, "applied")


def _apply_chart_op(slide, edit: PptxEdit) -> PptxEditResult:
    from contextifier.raw.opc import RawUnsupportedError

    charts = slide.charts
    if edit.chart is None or not (0 <= edit.chart < len(charts)):
        return PptxEditResult(edit.action, "not_found", "chart index out of range")
    chart = charts[edit.chart]
    try:
        if edit.action == "set_chart_title":
            if not (edit.title or "").strip():
                return PptxEditResult(edit.action, "invalid", "set_chart_title needs a title")
            chart.set_title(edit.title)
        else:  # set_chart_data
            if edit.categories is None or edit.series is None:
                return PptxEditResult(
                    edit.action, "invalid", "set_chart_data needs categories+series"
                )
            series: list[tuple[str | None, list]] = []
            for s in edit.series:
                if isinstance(s, dict):
                    series.append((s.get("name"), list(s.get("values") or [])))
                else:  # (name, values) tuple
                    series.append((s[0], list(s[1])))
            chart.set_data(categories=list(edit.categories), series=series)
    except RawUnsupportedError as exc:
        return PptxEditResult(edit.action, "unsupported", str(exc))
    except ValueError as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")
