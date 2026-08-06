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

import re
from collections.abc import Iterable
from dataclasses import dataclass

__all__ = [
    "PptxEdit",
    "PptxEditResult",
    "pptx_outline",
    "apply_pptx_edits",
    "find_shapes",
    "describe_ops",
    "OP_CATALOG",
    "VALID_PPTX_ACTIONS",
]

#: Self-describing operation catalog — the SINGLE SOURCE OF TRUTH for the
#: surgical PPTX interface. Each entry documents the action's address + payload
#: contract; validation, the planner prompt's op reference, and any external
#: tool schema are all derived from this (see :func:`describe_ops`). Keeping the
#: contract in one machine-readable place is what makes the framework
#: discoverable and drift-free.
OP_CATALOG: dict[str, dict] = {
    "set_text": {
        "summary": "Replace one paragraph's text in a text shape (in place).",
        "address": ["slide", "shape", "para?"],
        "payload": ["new_text", "old_text?"],
    },
    "set_shape_style": {
        "summary": "Restyle a text shape's font (color/size/bold/italic) and/or fill.",
        "address": ["slide", "shape", "para?"],
        "payload": ["color?", "size_pt?", "bold?", "italic?", "fill?"],
    },
    "set_shape_position": {
        "summary": "Move/resize a shape (inches).",
        "address": ["slide", "shape"],
        "payload": ["left?", "top?", "width?", "height?"],
    },
    "set_table_cell": {
        "summary": "Replace a table cell's text.",
        "address": ["slide", "shape", "row", "col"],
        "payload": ["new_text", "old_text?"],
    },
    "set_cell_style": {
        "summary": "Restyle a table cell (fill / font).",
        "address": ["slide", "shape", "row", "col"],
        "payload": ["fill?", "color?", "size_pt?", "bold?", "italic?"],
    },
    "insert_row": {
        "summary": "Insert a row into a table (clones the row above as template).",
        "address": ["slide", "shape"],
        "payload": ["at?"],
    },
    "delete_row": {
        "summary": "Delete a table row.",
        "address": ["slide", "shape", "row"],
        "payload": [],
    },
    "set_chart_data": {
        "summary": "Rewrite a chart's categories + series (classic charts).",
        "address": ["slide", "chart"],
        "payload": ["categories", "series"],
    },
    "set_chart_title": {
        "summary": "Set a chart's title.",
        "address": ["slide", "chart"],
        "payload": ["title"],
    },
    "add_textbox": {
        "summary": "Add a new text box at a position (inches) with text.",
        "address": ["slide"],
        "payload": ["new_text", "left", "top", "width", "height",
                    "color?", "size_pt?", "bold?", "italic?"],
    },
    "delete_shape": {
        "summary": "Delete a shape from a slide.",
        "address": ["slide", "shape"],
        "payload": [],
    },
    "duplicate_shape": {
        "summary": "Duplicate a shape (optionally offset by dx/dy inches).",
        "address": ["slide", "shape"],
        "payload": ["left?", "top?"],  # optional absolute position for the copy
    },
    "set_runs": {
        "summary": "Replace a paragraph with multiple independently-styled runs "
                   "(e.g. bold just one word).",
        "address": ["slide", "shape", "para?"],
        "payload": ["runs"],  # [{text, bold?, italic?, color?, size_pt?}, ...]
    },
    "insert_column": {
        "summary": "Insert a column into a table (clones the column to its left).",
        "address": ["slide", "shape"],
        "payload": ["at?"],
    },
    "delete_column": {
        "summary": "Delete a table column.",
        "address": ["slide", "shape", "col"],
        "payload": [],
    },
    "merge_cells": {
        "summary": "Merge a rectangular block of table cells.",
        "address": ["slide", "shape", "row", "col"],
        "payload": ["row2", "col2"],  # bottom-right of the block
    },
    "add_slide": {
        "summary": "Insert a NEW native slide from a layout, filling its "
                   "title/body placeholders (content fills the layout).",
        "address": ["after?"],  # insert after this 1-based slide (0 = start)
        "payload": ["layout?", "title?", "body?"],
    },
    "set_notes": {
        "summary": "Set a slide's speaker-notes text.",
        "address": ["slide"],
        "payload": ["new_text"],
    },
    "set_bullet": {
        "summary": "Set a paragraph's bullet (bullet / number / none).",
        "address": ["slide", "shape", "para?"],
        "payload": ["bullet?"],
    },
    "set_hyperlink": {
        "summary": "Attach a hyperlink to a text run or the whole shape.",
        "address": ["slide", "shape", "para?", "run?"],
        "payload": ["url"],
    },
    "set_z_order": {
        "summary": "Bring a shape to the front or send it to the back.",
        "address": ["slide", "shape"],
        "payload": ["order"],
    },
    "set_legend": {
        "summary": "Show/hide/position a chart legend (r/l/t/b/tr/none).",
        "address": ["slide", "chart"],
        "payload": ["position"],
    },
    "set_series_color": {
        "summary": "Set a chart series' fill color.",
        "address": ["slide", "chart"],
        "payload": ["series_index", "color"],
    },
    "set_theme_color": {
        "summary": "Swap a theme color deck-wide (accent1..6, dk1/2, lt1/2, hlink).",
        "address": [],
        "payload": ["theme_name", "color"],
    },
    "set_theme_font": {
        "summary": "Swap the major/minor theme font deck-wide.",
        "address": [],
        "payload": ["which", "typeface"],
    },
}

#: Every action the surgical engine understands (derived from the catalog).
VALID_PPTX_ACTIONS = tuple(OP_CATALOG)


def describe_ops() -> dict[str, dict]:
    """The machine-readable op catalog — lets any agent introspect exactly which
    surgical operations exist and each one's address + payload contract."""
    return {k: dict(v) for k, v in OP_CATALOG.items()}


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
        # One walk for all text paragraphs (avoids O(shapes^2) get_paragraphs).
        try:
            para_map = slide.paragraphs_by_shape()
        except Exception:
            para_map = {}
        for info in shapes:
            if info.id in tables:
                continue  # handled in the table pass below
            pos = _geometry_inches(info)
            if info.kind == "text":
                # Per-a:p text (NOT text.split("\n")) so the para index the
                # planner sees is the one set_text mutates, even with a literal
                # newline inside a run — read from the single-walk map.
                paras = para_map.get(info.id)
                if paras is None:
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
    slide: int | None = None         # 1-based slide number (None for add_slide)
    shape: int | None = None         # cNvPr id (text / table ops)
    para: int | None = None          # paragraph index (None = whole shape / para 0)
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
    # set_shape_position / add_textbox — inches
    left: float | None = None
    top: float | None = None
    width: float | None = None
    height: float | None = None
    # set_runs
    runs: list | None = None         # [{text, bold?, italic?, color?, size_pt?}, ...]
    # merge_cells — bottom-right of the block
    row2: int | None = None
    col2: int | None = None
    # add_slide
    after: int | None = None         # insert after this 1-based slide (0 = start)
    layout: object = None            # layout index (int) or a name/type hint (str)
    body: str | None = None          # body placeholder text (lines split on \n)
    # set_notes / set_bullet / set_hyperlink / set_z_order / chart depth / theme
    url: str | None = None
    bullet: str | None = None        # "bullet" | "number" | "none"
    order: str | None = None         # "front" | "back"
    run: int | None = None           # run index for set_hyperlink
    position: str | None = None      # legend position
    series_index: int | None = None  # series ordinal for set_series_color
    theme_name: str | None = None    # theme color slot
    typeface: str | None = None      # theme font
    which: str | None = None


@dataclass
class PptxEditResult:
    action: str
    status: str  # applied | stale | not_found | invalid | unsupported
    message: str = ""


def _normalize(text: str) -> str:
    return " ".join((text or "").split())


def _as_number(v):
    """Coerce a chart value to a FINITE float; None stays None (a gap in the
    series). A non-numeric or non-finite value raises ValueError → the op is
    reported ``invalid`` (never writes ``nan``/``inf`` into a numCache).

    Only spaces and thousands separators around whole groups are stripped —
    an ambiguous ``"1,5"`` (European decimal) is NOT silently turned into 15."""
    import math

    if v is None:
        return None
    if isinstance(v, bool):
        raise ValueError(f"chart value {v!r} is not numeric")
    if isinstance(v, (int, float)):
        f = float(v)
    else:
        s = str(v).strip()
        # Strip thousands separators only when unambiguous (3-digit groups).
        if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", s):
            s = s.replace(",", "")
        f = float(s)
    if not math.isfinite(f):
        raise ValueError(f"chart value {v!r} is not finite")
    return f


def _safe_int(v, default: int = -1) -> int:
    if isinstance(v, bool):  # bool is an int subclass — never a valid address
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _opt_int(v):
    """int(v) or None — keeps None as None; used to normalize address fields so
    a string like '0' from any caller behaves like the int 0. A bool is NOT a
    valid address, so it maps to None (rejected downstream)."""
    if v is None or isinstance(v, bool):
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _row_sort_key(edit: PptxEdit) -> tuple:
    """Order ops within a table so addresses stay valid.

    Applied with ``reverse=True``: higher row index first (so an insert/delete
    at a high index doesn't shift a lower address before we reach it). The
    trailing flag breaks ties at the SAME index so a structural op
    (insert_row/delete_row) always runs BEFORE a content op
    (set_table_cell/set_cell_style) — i.e. ``insert_row at=1`` then
    ``set_table_cell row=1`` deterministically populates the NEW row,
    regardless of the order the planner emitted them. All fields are coerced so
    a stray string address can never crash the sort."""
    if edit.action in ("insert_column", "delete_column"):
        idx = edit.at if edit.at is not None else edit.col
    else:
        idx = edit.at if edit.at is not None else edit.row
    structural = 1 if edit.action in (
        "insert_row", "delete_row", "insert_column", "delete_column"
    ) else 0
    return (_safe_int(edit.slide), _safe_int(edit.shape), _safe_int(idx), structural)


def apply_pptx_edits(
    content: bytes,
    edits: Iterable[PptxEdit],
    *,
    dry_run: bool = False,
    atomic: bool = False,
) -> tuple[bytes, list[PptxEditResult]]:
    """Apply surgical edits losslessly via contextifier's raw layer.

    Only the slide parts an edit touches are rewritten; every other part of the
    package stays byte-identical. Per-edit soft failures never abort the batch —
    each op yields an ``applied | stale | not_found | invalid | unsupported``
    result, reported in caller order.

    * ``dry_run`` — run every op for real against an in-memory copy (so the
      statuses are accurate), but return the ORIGINAL bytes: validate a plan
      without changing the file.
    * ``atomic`` — all-or-nothing: if any op is not ``applied``, ship nothing
      (return the original bytes) so the deck is never left half-edited.
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

    # Row/column insert/delete shift indices within one table; run them
    # highest-index-first so each op's original address stays valid. Non-shifting
    # ops are order-independent (shape id / chart ordinal don't shift).
    ordered = sorted(enumerate(edit_list), key=lambda pair: _row_sort_key(pair[1]), reverse=True)
    for index, edit in ordered:
        results[index] = _apply_one(raw, edit)

    final = [r for r in results if r is not None]
    all_ok = all(r.status == "applied" for r in final)
    # Dry run ships nothing; atomic ships nothing unless every op applied.
    if dry_run or (atomic and not all_ok):
        return content, final
    return raw.to_bytes(), final


def find_shapes(
    content: bytes,
    *,
    text: str | None = None,
    kind: str | None = None,
    name: str | None = None,
    slide: int | None = None,
) -> list[dict]:
    """Query the deck for shapes matching criteria — a selection interface so an
    agent can target edits without eyeballing the whole outline.

    Filters (all optional, ANDed): ``text`` (case-insensitive substring of the
    shape's text), ``kind`` (text/table/chart/picture/…), ``name`` (substring of
    the shape name), ``slide`` (1-based). Returns matching outline entries
    (``slide``, ``shape``/``chart``, ``kind``, ``name``, ``text`` …).
    """
    tl = text.lower() if text else None
    nl = name.lower() if name else None
    out: list[dict] = []
    for e in pptx_outline(content):
        if slide is not None and e.get("slide") != slide:
            continue
        if kind is not None and e.get("kind") != kind:
            continue
        if nl is not None and nl not in (e.get("name") or "").lower():
            continue
        if tl is not None:
            blob = (e.get("text") or e.get("title") or "").lower()
            if tl not in blob:
                continue
        out.append(e)
    return out


def _slide_of(raw, edit: PptxEdit):
    """Resolve the 1-based ``edit.slide`` to a RawSlide, or None if invalid.

    Coerces defensively so a non-int slide never raises out of the batch."""
    slides = raw.slides
    n = _safe_int(edit.slide, 0)
    if not (1 <= n <= len(slides)):
        return None
    return slides[n - 1]


def _apply_one(raw, edit: PptxEdit) -> PptxEditResult:
    if edit.action not in VALID_PPTX_ACTIONS:
        return PptxEditResult(edit.action, "invalid", f"unknown action {edit.action!r}")

    # Document-level ops (no single slide) are handled before slide resolution.
    if edit.action == "add_slide":
        return _apply_add_slide(raw, edit)
    if edit.action in ("set_theme_color", "set_theme_font"):
        return _apply_theme_op(raw, edit)

    slide = _slide_of(raw, edit)
    if slide is None:
        return PptxEditResult(edit.action, "not_found", "slide index out of range")

    # Normalize address fields so a string like "0" behaves like int 0 for any
    # caller (the LLM boundary also coerces; this is defense in depth).
    edit.shape = _opt_int(edit.shape)
    edit.para = _opt_int(edit.para)
    edit.row = _opt_int(edit.row)
    edit.col = _opt_int(edit.col)
    edit.at = _opt_int(edit.at)
    edit.chart = _opt_int(edit.chart)
    edit.row2 = _opt_int(edit.row2)
    edit.col2 = _opt_int(edit.col2)
    edit.run = _opt_int(edit.run)
    edit.series_index = _opt_int(edit.series_index)

    try:
        if edit.action == "set_text":
            return _apply_set_text(slide, edit)
        if edit.action == "set_runs":
            return _apply_set_runs(slide, edit)
        if edit.action == "set_shape_style":
            return _apply_set_shape_style(slide, edit)
        if edit.action == "set_shape_position":
            return _apply_set_shape_position(slide, edit)
        if edit.action == "add_textbox":
            return _apply_add_textbox(slide, edit)
        if edit.action == "delete_shape":
            return _apply_delete_shape(slide, edit)
        if edit.action == "duplicate_shape":
            return _apply_duplicate_shape(slide, edit)
        if edit.action == "set_table_cell":
            return _apply_set_table_cell(slide, edit)
        if edit.action == "set_cell_style":
            return _apply_set_cell_style(slide, edit)
        if edit.action in ("insert_row", "delete_row", "insert_column", "delete_column"):
            return _apply_row_op(slide, edit)
        if edit.action == "merge_cells":
            return _apply_merge_cells(slide, edit)
        if edit.action == "set_notes":
            slide.set_notes(edit.new_text or "")
            return PptxEditResult(edit.action, "applied")
        if edit.action == "set_bullet":
            return _apply_set_bullet(slide, edit)
        if edit.action == "set_hyperlink":
            return _apply_set_hyperlink(slide, edit)
        if edit.action == "set_z_order":
            return _apply_set_z_order(slide, edit)
        if edit.action in ("set_legend", "set_series_color"):
            return _apply_chart_depth(slide, edit)
        if edit.action in ("set_chart_data", "set_chart_title"):
            return _apply_chart_op(slide, edit)
    except Exception as exc:  # last-resort: one bad op never kills the batch
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "invalid", f"unhandled action {edit.action!r}")


def _apply_set_text(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_text needs a shape id")
    try:
        paras = slide.get_paragraphs(edit.shape)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except ValueError:
        return PptxEditResult(edit.action, "invalid", "shape has no editable text")
    para = edit.para or 0  # set_text targets a single paragraph (None → first)
    if edit.old_text is not None:
        current = paras[para] if 0 <= para < len(paras) else ""
        if _normalize(current) != _normalize(edit.old_text):
            return PptxEditResult(edit.action, "stale", "paragraph text changed; refresh")
    try:
        slide.set_text(edit.shape, edit.new_text, para=para)
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "paragraph index out of range")
    return PptxEditResult(edit.action, "applied")


#: placeholder types that hold the slide title / the main body.
_TITLE_PH = ("title", "ctrTitle")
_BODY_PH = ("body", "subTitle", "obj")


def _pick_layout(layouts: list[dict], hint) -> int:
    """Resolve a layout index from an int index or a name/type hint, defaulting
    to a title+body layout ('Title and Content' / obj) so a plain add_slide
    still fills a real layout."""
    if isinstance(hint, bool):
        hint = None
    if isinstance(hint, int):
        return hint if 0 <= hint < len(layouts) else _pick_layout(layouts, None)
    if isinstance(hint, str) and hint.strip():
        h = hint.strip().lower()
        for lay in layouts:  # match on name or type
            if h in (lay.get("name") or "").lower() or h == (lay.get("type") or "").lower():
                return lay["index"]
    # default: a layout with both a title and a body placeholder
    for lay in layouts:
        types = {p.get("type") for p in lay.get("placeholders", [])}
        if types & set(_TITLE_PH) and types & set(_BODY_PH):
            return lay["index"]
    return 0


def _apply_add_slide(raw, edit: PptxEdit) -> PptxEditResult:
    layouts = raw.layouts
    if not layouts:
        return PptxEditResult(edit.action, "unsupported", "deck has no slide layouts")
    layout_index = _pick_layout(layouts, edit.layout)
    # `after` is 1-based (0 = start); contextifier add_slide uses a 0-based
    # insertion position = number of slides before the new one.
    at = None if edit.after is None else max(0, _safe_int(edit.after, len(raw.slides)))
    try:
        new_slide = raw.add_slide(layout_index, at=at)
    except (IndexError, ValueError) as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    # fill title / body placeholders — the new slide's text shapes are created
    # in the layout's placeholder order, so map them back to their ph types.
    title_id = body_id = None
    layout_phs = [p for p in layouts[layout_index]["placeholders"]]
    text_shapes = [s for s in new_slide.shapes if s.kind == "text"]
    for info, ph in zip(text_shapes, layout_phs, strict=False):
        t = ph.get("type")
        if t in _TITLE_PH and title_id is None:
            title_id = info.id
        elif t in _BODY_PH and body_id is None:
            body_id = info.id
    if edit.title and title_id is not None:
        new_slide.set_text(title_id, edit.title)
    elif edit.title and text_shapes:
        new_slide.set_text(text_shapes[0].id, edit.title)
    if edit.body and body_id is not None:
        new_slide.set_paragraphs(body_id, edit.body.split("\n"))
    elif edit.body and edit.new_text is None and len(text_shapes) > 1 and body_id is None:
        new_slide.set_paragraphs(text_shapes[1].id, edit.body.split("\n"))
    return PptxEditResult(edit.action, "applied")


def _apply_set_runs(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_runs needs a shape id")
    if not edit.runs or not isinstance(edit.runs, list):
        return PptxEditResult(edit.action, "invalid", "set_runs needs a non-empty runs list")
    specs = []
    for r in edit.runs:
        if not isinstance(r, dict):
            return PptxEditResult(edit.action, "invalid", "each run must be an object with text")
        specs.append({
            "text": str(r.get("text", "")),
            "color": r.get("color"),
            "size_pt": r.get("size_pt"),
            "bold": r.get("bold"),
            "italic": r.get("italic"),
        })
    try:
        slide.set_runs(edit.shape, edit.para or 0, specs)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "paragraph index out of range")
    except ValueError:
        return PptxEditResult(edit.action, "invalid", "shape has no editable text")
    return PptxEditResult(edit.action, "applied")


def _apply_add_textbox(slide, edit: PptxEdit) -> PptxEditResult:
    for f in ("left", "top", "width", "height"):
        if getattr(edit, f) is None:
            return PptxEditResult(edit.action, "invalid", f"add_textbox needs {f} (inches)")

    def _emu(v):
        return int(round(v * _EMU_PER_INCH))

    slide.add_textbox(
        edit.new_text,
        left=_emu(edit.left), top=_emu(edit.top),
        width=_emu(edit.width), height=_emu(edit.height),
        color=edit.color, size_pt=edit.size_pt, bold=edit.bold, italic=edit.italic,
    )
    return PptxEditResult(edit.action, "applied")


def _apply_delete_shape(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "delete_shape needs a shape id")
    try:
        slide.delete_shape(edit.shape)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    return PptxEditResult(edit.action, "applied")


def _apply_duplicate_shape(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "duplicate_shape needs a shape id")

    def _emu(v):
        return int(round(v * _EMU_PER_INCH)) if v is not None else None

    try:
        slide.duplicate_shape(edit.shape, left=_emu(edit.left), top=_emu(edit.top))
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    return PptxEditResult(edit.action, "applied")


def _apply_merge_cells(slide, edit: PptxEdit) -> PptxEditResult:
    table = _table_by_shape(slide, edit.shape)
    if table is None:
        return PptxEditResult(edit.action, "not_found", "no table with that shape id")
    if None in (edit.row, edit.col, edit.row2, edit.col2):
        return PptxEditResult(edit.action, "invalid", "merge_cells needs row,col,row2,col2")
    try:
        table.merge_cells(edit.row, edit.col, edit.row2, edit.col2)
    except IndexError as exc:
        return PptxEditResult(edit.action, "not_found", str(exc))
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
                para=edit.para,  # None = whole shape; 0 = first paragraph only
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
    coords = (edit.left, edit.top, edit.width, edit.height)
    if all(v is None for v in coords):
        return PptxEditResult(edit.action, "invalid", "set_shape_position needs a coordinate")

    # A shape that inherits its placement from a placeholder has no explicit
    # a:xfrm; creating one with only some coordinates would zero the rest and
    # shrink the shape to nothing. Require all four in that case.
    info = next((s for s in slide.shapes if s.id == edit.shape), None)
    if info is None:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    inherits_placement = getattr(info, "left", None) is None
    if inherits_placement and any(v is None for v in coords):
        return PptxEditResult(
            edit.action, "invalid",
            "this shape inherits its placement — provide left, top, width AND height",
        )

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


def _apply_theme_op(raw, edit: PptxEdit) -> PptxEditResult:
    try:
        if edit.action == "set_theme_color":
            if not (edit.theme_name and edit.color):
                return PptxEditResult(
                    edit.action, "invalid", "set_theme_color needs theme_name+color"
                )
            raw.set_theme_color(str(edit.theme_name), str(edit.color))
        else:  # set_theme_font
            if not (edit.which and edit.typeface):
                return PptxEditResult(edit.action, "invalid", "set_theme_font needs which+typeface")
            raw.set_theme_font(str(edit.which), str(edit.typeface))
    except ValueError as exc:
        return PptxEditResult(edit.action, "not_found", str(exc))
    return PptxEditResult(edit.action, "applied")


def _apply_set_bullet(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_bullet needs a shape id")
    style = (edit.bullet or "bullet").lower()
    if style not in ("bullet", "number", "none"):
        return PptxEditResult(edit.action, "invalid", "bullet must be bullet|number|none")
    try:
        slide.set_paragraph_bullet(edit.shape, edit.para or 0, style=style)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "paragraph index out of range")
    except ValueError:
        return PptxEditResult(edit.action, "invalid", "shape has no editable text")
    return PptxEditResult(edit.action, "applied")


def _apply_set_hyperlink(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_hyperlink needs a shape id")
    if not (edit.url or "").strip():
        return PptxEditResult(edit.action, "invalid", "set_hyperlink needs a url")
    try:
        slide.set_hyperlink(edit.shape, str(edit.url), para=edit.para, run=edit.run)
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except IndexError:
        return PptxEditResult(edit.action, "not_found", "paragraph/run index out of range")
    except ValueError as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")


def _apply_set_z_order(slide, edit: PptxEdit) -> PptxEditResult:
    if edit.shape is None:
        return PptxEditResult(edit.action, "invalid", "set_z_order needs a shape id")
    order = (edit.order or "front").lower()
    if order not in ("front", "back"):
        return PptxEditResult(edit.action, "invalid", "order must be front|back")
    try:
        slide.set_z_order(edit.shape, to_front=(order == "front"))
    except KeyError:
        return PptxEditResult(edit.action, "not_found", "no shape with that id")
    except ValueError as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")


def _apply_chart_depth(slide, edit: PptxEdit) -> PptxEditResult:
    from contextifier.raw.opc import RawUnsupportedError

    charts = slide.charts
    if edit.chart is None or not (0 <= edit.chart < len(charts)):
        return PptxEditResult(edit.action, "not_found", "chart index out of range")
    chart = charts[edit.chart]
    try:
        if edit.action == "set_legend":
            chart.set_legend(edit.position or "r")
        else:  # set_series_color
            if edit.series_index is None or not (edit.color or "").strip():
                return PptxEditResult(
                    edit.action, "invalid", "set_series_color needs series_index+color"
                )
            chart.set_series_color(edit.series_index, str(edit.color))
    except RawUnsupportedError as exc:
        return PptxEditResult(edit.action, "unsupported", str(exc))
    except (IndexError, ValueError) as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")


def _table_by_shape(slide, shape_id):
    if shape_id is None:
        return None
    sid = _safe_int(shape_id, None)
    for table in slide.tables:
        if table.shape_id == sid:
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
            table.insert_row(edit.at if edit.at is not None else table.n_rows)
        elif edit.action == "delete_row":
            if edit.row is None:
                return PptxEditResult(edit.action, "invalid", "delete_row needs row")
            table.delete_row(edit.row)
        elif edit.action == "insert_column":
            table.insert_column(edit.at if edit.at is not None else table.n_cols)
        else:  # delete_column
            if edit.col is None:
                return PptxEditResult(edit.action, "invalid", "delete_column needs col")
            table.delete_column(edit.col)
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
                name, values = (
                    (s.get("name"), s.get("values") or []) if isinstance(s, dict)
                    else (s[0], s[1])
                )
                series.append((name, [_as_number(v) for v in values]))
            chart.set_data(
                categories=["" if c is None else str(c) for c in edit.categories],
                series=series,
            )
    except RawUnsupportedError as exc:
        return PptxEditResult(edit.action, "unsupported", str(exc))
    except ValueError as exc:
        return PptxEditResult(edit.action, "invalid", str(exc))
    return PptxEditResult(edit.action, "applied")
