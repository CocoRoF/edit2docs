"""Chat-edit orchestrator for DOCX / XLSX (one turn = one call).

Same contract as the PPTX chat editor (tools/edit_deck.py): a planner LLM
turns the instruction + a structural outline into a minimal operation
list plus a chat reply; the deterministic engines apply the operations so
untouched content survives byte-identical. Question-only turns answer
without changing the file. Plan-block failures retry once and then admit
failure in the reply instead of promising changes.
"""

from __future__ import annotations

import re
import time
from typing import Literal

from pydantic import Field

from ..config import resolve_model
from ..documents.docx_engine import DocxEdit, apply_docx_edits, docx_outline
from ..documents.pptx_engine import (
    VALID_PPTX_ACTIONS,
    PptxEdit,
    apply_pptx_edits,
    pptx_outline,
)
from ..documents.xlsx_engine import XlsxEdit, apply_xlsx_edits, xlsx_outline
from ..llm import DEFAULT_MODEL, AnthropicClient, build_output_lang_directive, load_prompt
from ._edit_events import op_event_vars, op_summary, plan_event_vars
from ._reply_texts import reply_text
from .edit_deck import ChatTurn, _cost_from_usage, _parse_plan
from .generate_deck import EventCallback, StageEvent, _emit, _merge_cost
from .types import (
    DEFAULT_LANG,
    CostBreakdown,
    LangCode,
    ToolRequest,
    ToolResponse,
    WarningEntry,
)

DocFormat = Literal["docx", "xlsx", "pptx"]

_PLANNER_ROLE = {
    "docx": "doc-editor-planner",
    "xlsx": "sheet-editor-planner",
    "pptx": "pptx-editor-planner",
}
_MAX_OPERATIONS = 30

# Retry reminder for a plan-missing first response. Passed as the LLM call's
# ``user_suffix`` (the volatile tail) so the stable ``user`` prefix — the
# outline + sources + history, unbounded on a large docx — stays
# byte-identical between the first call and the retry and is READ from cache
# on the retry instead of re-sent at full price.
_PLAN_REMINDER = (
    "\n\n# REMINDER\nYour previous answer was missing the "
    "```edit_plan fenced block. Respond again following the "
    "output format EXACTLY (operations: [] only if truly no "
    "change is needed)."
)

# Above this outline size (characters), _outline_context emits a WINDOWED
# outline (headings skeleton + a paragraph window around a relevance anchor)
# instead of the full, unbounded paragraph listing — which is re-sent every
# turn and grows linearly with the document (300 paras ≈ 19k tokens).
_OUTLINE_CHAR_BUDGET = 40000
# When windowing, how many full paragraph lines to keep at the head and tail
# (when there is no in-text anchor), and the half-width of the window we
# center on an anchor paragraph.
_WINDOW_HEAD = 40
_WINDOW_TAIL = 20
_ANCHOR_RADIUS = 30


class EditDocRequest(ToolRequest):
    content: bytes
    fmt: DocFormat
    instruction: str = Field(..., min_length=1)
    chat_history: list[ChatTurn] = Field(default_factory=list)
    sources_markdown: list[str] = Field(default_factory=list)
    lang: LangCode = DEFAULT_LANG
    model: str = DEFAULT_MODEL
    anthropic_api_key: str = Field(..., description="BYOK; never persisted.")


class EditDocResponse(ToolResponse):
    content: bytes
    changed: bool
    reply: str
    operations: list[dict] = Field(default_factory=list)
    cost: CostBreakdown
    warnings: list[WarningEntry] = Field(default_factory=list)
    # PPTX only: the surgical planner judged this turn needs the SVG
    # redesign/generation path (creating new slides or a full visual
    # redesign that can't be expressed as targeted text/table/chart edits).
    # The executor falls back to tools.edit_deck when this is set.
    needs_svg: bool = False


#: PPTX planner sentinels that mean "hand this turn to the SVG generator".
#: (add_slide is now a NATIVE, layout-aware op — no longer an SVG sentinel.)
_PPTX_GENERATIVE_ACTIONS = {"redesign"}


async def edit_document(
    req: EditDocRequest, *, on_event: EventCallback = None
) -> EditDocResponse:
    started = time.perf_counter()
    warnings: list[WarningEntry] = []
    cost = CostBreakdown()

    await _emit(
        on_event,
        StageEvent(stage="planning_edits", progress=0.2, message_key="stages.planning_edits"),
    )
    outline_text = _outline_context(req, warnings)
    client = AnthropicClient(api_key=req.anthropic_api_key, model=req.model)
    system = (
        build_output_lang_directive(req.lang)
        + "\n\n"
        + load_prompt(_PLANNER_ROLE[req.fmt])
    )
    user = _build_user_message(req, outline_text)

    planner_model = resolve_model("planner", req.model)
    result = await client.complete(
        system_prompt=system,
        user_message=user,
        max_output_tokens=16384,
        cache_system=True,
        model=planner_model,
    )
    cost = _merge_cost(cost, _cost_from_usage(result.usage))
    reply, raw_ops, plan_missing = _parse_plan(result.text, warnings, lang=req.lang)

    if plan_missing:
        # Retry re-sends the SAME ``user`` (cache read) with the reminder in
        # ``user_suffix`` — no re-paying for the outline/sources/history.
        retry = await client.complete(
            system_prompt=system,
            user_message=user,
            user_suffix=_PLAN_REMINDER,
            max_output_tokens=16384,
            cache_system=True,
            model=planner_model,
        )
        cost = _merge_cost(cost, _cost_from_usage(retry.usage))
        reply, raw_ops, plan_missing = _parse_plan(retry.text, warnings, lang=req.lang)
        if plan_missing:
            reply = reply.rstrip() + reply_text("plan_failed", req.lang)

    # PPTX: if the surgical planner decided the request needs new/redesigned
    # slides (beyond targeted text/table/chart edits), signal the executor to
    # hand the turn to the SVG generator (tools.edit_deck). Any surgical ops the
    # planner co-emitted are applied FIRST (so they aren't lost) and the
    # surgically-edited deck is what the SVG generator continues from.
    if req.fmt == "pptx" and any(
        isinstance(o, dict) and o.get("action") in _PPTX_GENERATIVE_ACTIONS
        for o in raw_ops
    ):
        surgical = [
            o for o in raw_ops
            if isinstance(o, dict) and o.get("action") in VALID_PPTX_ACTIONS
        ][:_MAX_OPERATIONS]  # same cap as the normal path
        new_content, applied_ops = req.content, []
        if surgical:
            new_content, applied_ops, op_warnings, _ = _apply("pptx", req.content, surgical)
            warnings.extend(op_warnings)
        await _emit(on_event, StageEvent(stage="done", progress=1.0, message_key="stages.done"))
        return EditDocResponse(
            content=new_content, changed=bool(applied_ops), reply=reply,
            operations=applied_ops,
            cost=CostBreakdown(
                input_tokens=cost.input_tokens, output_tokens=cost.output_tokens,
                cache_read_tokens=cost.cache_read_tokens,
                cache_write_tokens=cost.cache_write_tokens,
                duration_seconds=time.perf_counter() - started,
            ),
            warnings=warnings, needs_svg=True,
        )

    if len(raw_ops) > _MAX_OPERATIONS:
        warnings.append(
            WarningEntry(
                code="edit_plan_truncated",
                message=f"{len(raw_ops)} operations planned; applying first {_MAX_OPERATIONS}.",
                detail={"emitted": len(raw_ops), "cap": _MAX_OPERATIONS},
            )
        )
        reply = reply.rstrip() + reply_text(
            "plan_truncated", req.lang, emitted=len(raw_ops), cap=_MAX_OPERATIONS
        )
        raw_ops = raw_ops[:_MAX_OPERATIONS]

    applied_ops: list[dict] = []
    new_content = req.content
    if raw_ops:
        # Announce the plan, then stream each op's result as it's applied.
        valid_ops = [
            op for op in raw_ops
            if isinstance(op, dict) and op.get("action") in _VALID_ACTIONS[req.fmt]
        ]
        await _emit(
            on_event,
            StageEvent(
                stage="editing_slides", progress=0.6,
                message_key="stages.editing_slides",
                message_vars=plan_event_vars(req.fmt, valid_ops, lang=req.lang),
            ),
        )
        new_content, applied_ops, op_warnings, op_results = _apply(
            req.fmt, req.content, raw_ops
        )
        warnings.extend(op_warnings)
        total = len(op_results)
        for i, (op, status) in enumerate(op_results):
            await _emit(
                on_event,
                StageEvent(
                    stage="applying_edits", progress=0.85,
                    message_key="stages.applying_edits",
                    message_vars=op_event_vars(
                        op_summary(req.fmt, op, index=i, total=total, lang=req.lang),
                        phase="done", status=status,
                    ),
                ),
            )

    changed = bool(applied_ops)
    await _emit(on_event, StageEvent(stage="done", progress=1.0, message_key="stages.done"))
    return EditDocResponse(
        content=new_content if changed else req.content,
        changed=changed,
        reply=reply,
        operations=applied_ops,
        cost=CostBreakdown(
            input_tokens=cost.input_tokens,
            output_tokens=cost.output_tokens,
            cache_read_tokens=cost.cache_read_tokens,
            cache_write_tokens=cost.cache_write_tokens,
            duration_seconds=time.perf_counter() - started,
        ),
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# Context assembly
# ---------------------------------------------------------------------------


def _outline_context(req: EditDocRequest, warnings: list[WarningEntry]) -> str:
    if req.fmt == "docx":
        return _docx_outline_context(req, warnings)
    if req.fmt == "pptx":
        return _pptx_outline_context(req, warnings)

    # xlsx is already sample-bounded (sample_rows caps the per-sheet body),
    # so its outline size is independent of the workbook's row count — no
    # windowing needed here.
    lines = ["# Workbook outline"]
    for sheet in xlsx_outline(req.content, sample_rows=12)["sheets"]:
        lines.append(
            f"## sheet {sheet['name']!r} — {sheet['rows']} rows x {sheet['columns']} cols"
        )
        for r, row in enumerate(sheet["sample"], start=1):
            rendered = ", ".join(
                "" if v is None else str(v) for v in row
            )
            lines.append(f"- row {r}: {rendered[:200]}")
    return "\n".join(lines)


def _pos_str(pos: dict | None) -> str:
    """Compact geometry hint for a shape line (inches)."""
    if not pos:
        return ""
    return (
        f" @({pos['left']},{pos['top']} {pos['width']}x{pos['height']}in)"
    )


def _pptx_entry_line(e: dict) -> str:
    """Render one pptx_outline entry in the address format the planner uses."""
    kind = e.get("kind")
    if "table" in e and "row" in e:  # a table cell
        return f"  - table {e['table']} cell ({e['row']},{e['col']}): {e['text'][:100]}"
    if kind == "text":
        name = f" \"{e['name']}\"" if e.get("name") else ""
        return (
            f"- shape {e['shape']}{name} para {e['para']}"
            f"{_pos_str(e.get('pos'))}: {e['text'][:140]}"
        )
    if kind == "table":
        return f"- table {e['shape']} ({e['rows']}x{e['cols']} — edit cells by (row,col))"
    if kind == "chart":
        ser = "; ".join(
            f"{s['name']}=[{', '.join(str(v) for v in s['values'][:8])}]"
            for s in (e.get("series") or [])[:4]
        )
        cats = ", ".join(
            str(c) for c in ((e.get("series") or [{}])[0].get("categories") or [])[:8]
        )
        title = f" \"{e['title']}\"" if e.get("title") else ""
        return (
            f"- chart {e['chart']} [{e.get('chart_type')}]{title}: "
            f"categories=[{cats}] series: {ser}"
        )
    name = f" \"{e['name']}\"" if e.get("name") else ""
    return f"- shape {e['shape']} [{kind}]{name}{_pos_str(e.get('pos'))}"


_SLIDE_REF = re.compile(r"\bslides?\.?\s*#?\s*(\d+)", re.IGNORECASE)
_SLIDE_REF_KO = re.compile(r"(\d+)\s*(?:번째?|장)\s*(?:슬라이드|장|페이지)?")


def _pptx_anchor_slides(
    req: EditDocRequest, max_slide: int, per_slide: dict[int, list[dict]] | None = None
) -> set[int]:
    """1-based slide numbers the instruction/chat points at (for windowing).

    Two signals, mirroring the docx anchor: an explicit slide number ("slide 3",
    "3번 슬라이드") AND quoted / distinctive text that matches a slide's content
    (so "change the title on the Revenue slide" anchors the Revenue slide even
    without a number)."""
    text = "\n".join(
        [req.instruction] + [t.content for t in req.chat_history[-6:]]
    )
    out: set[int] = set()
    for m in list(_SLIDE_REF.finditer(text)) + list(_SLIDE_REF_KO.finditer(text)):
        n = int(m.group(1))
        if 1 <= n <= max_slide:
            out.add(n)
    # Content anchor: quoted phrases (and, failing that, salient words) matched
    # against each slide's text.
    if per_slide:
        quotes = [q.strip().lower() for q in _QUOTED.findall(text) if len(q.strip()) >= 3]
        if quotes:
            for n, entries in per_slide.items():
                blob = " ".join(
                    str(e.get("text") or e.get("title") or "") for e in entries
                ).lower()
                if any(q in blob for q in quotes):
                    out.add(n)
    return out


def _pptx_outline_context(req: EditDocRequest, warnings: list[WarningEntry]) -> str:
    """Format the addressable slide model into the planner's address lines.

    Small decks are sent in full. A large deck (over the char budget — many
    slides or big tables, re-sent every turn) is WINDOWED: slides the
    instruction references (± a neighbour) keep full detail; the rest collapse
    to a one-line summary so every slide stays addressable and the planner can
    ask to expand one."""
    entries = pptx_outline(req.content)
    per_slide: dict[int, list[dict]] = {}
    for e in entries:
        per_slide.setdefault(e["slide"], []).append(e)

    header = [
        "# Deck outline (slide / shape / table / chart addresses)",
        "# Edit in place — reference the shape/table/chart by the ids shown.",
    ]

    def _slide_block(n: int) -> list[str]:
        return [f"## slide {n}"] + [_pptx_entry_line(e) for e in per_slide[n]]

    full = header + [ln for n in sorted(per_slide) for ln in _slide_block(n)]
    full_text = "\n".join(full)
    if len(full_text) <= _OUTLINE_CHAR_BUDGET:
        return full_text

    # Windowed: keep anchored slides (± 1) in full, summarize the rest.
    max_slide = max(per_slide) if per_slide else 0
    anchors = _pptx_anchor_slides(req, max_slide, per_slide)
    keep = set()
    for a in (anchors or {1, 2}):
        keep.update({a - 1, a, a + 1})

    out = header + [
        "# WINDOWED (large deck): slides you referenced are shown in full; "
        "others are summarized — ask to see a specific slide for full detail.",
    ]
    for n in sorted(per_slide):
        if n in keep:
            out.extend(_slide_block(n))
        else:
            counts: dict[str, int] = {}
            for e in per_slide[n]:
                k = "cell" if ("table" in e and "row" in e) else e.get("kind", "shape")
                counts[k] = counts.get(k, 0) + 1
            summary = ", ".join(f"{v} {k}" for k, v in counts.items())
            out.append(f"## slide {n} — ({summary}; ask to expand)")
    warnings.append(
        WarningEntry(
            code="pptx_outline_windowed",
            message=(
                f"Deck outline exceeded {_OUTLINE_CHAR_BUDGET} chars; sent a "
                "windowed view (referenced slides in full, others summarized)."
            ),
            detail={"slides": max_slide, "shown_full": sorted(keep & set(per_slide))},
        )
    )
    return "\n".join(out)


def _outline_line(entry: dict) -> str:
    """Render one docx_outline entry in the exact address format the planner's
    edit ops resolve against (``para N [style]: text`` / table / chart)."""
    if "para" in entry:
        return f"- para {entry['para']} [{entry['style']}]: {entry['text'][:160]}"
    if "table" in entry:
        return (
            f"- table {entry['table']} cell ({entry['row']},{entry['col']}): "
            f"{entry['text'][:120]}"
        )
    # Read-only chart summary appended by docx_outline.
    return f"- chart {entry.get('chart')} [{entry.get('kind')}]: {entry.get('title') or ''}"


def _is_heading(style: str | None) -> bool:
    s = (style or "").strip().lower()
    return s.startswith("heading") or s in ("title", "subtitle")


def _docx_outline_context(req: EditDocRequest, warnings: list[WarningEntry]) -> str:
    """Docx outline, windowed when it would blow the token budget.

    Small documents (full outline under ``_OUTLINE_CHAR_BUDGET``) are sent
    verbatim — no behaviour change. A large document (unbounded, linear in
    paragraph count, and re-sent every turn) instead gets a WINDOWED view:
    the full heading/table skeleton is always present, plus a window of full
    paragraph lines around a relevance anchor. Omitted runs collapse into a
    marker, and a ``WarningEntry("outline_windowed", ...)`` is emitted. Every
    rendered line keeps the exact ``para N [style]`` address format so edit
    ops still resolve.
    """
    entries = docx_outline(req.content)
    full = "\n".join(
        ["# Document outline (paragraph addresses)"]
        + [_outline_line(e) for e in entries]
    )
    if len(full) <= _OUTLINE_CHAR_BUDGET:
        return full

    # Ordinal of each paragraph within the paragraph sequence (para indices
    # themselves are non-contiguous — empty paragraphs are skipped upstream).
    para_ordinal = {
        e["para"]: j
        for j, e in enumerate(e for e in entries if "para" in e)
    }
    total_paras = len(para_ordinal)

    anchors = _anchor_para_indices(req, entries)
    keep_ordinals: set[int] = set()
    if anchors:
        for para_idx in anchors:
            o = para_ordinal.get(para_idx)
            if o is not None:
                keep_ordinals.update(range(o - _ANCHOR_RADIUS, o + _ANCHOR_RADIUS + 1))
    else:
        keep_ordinals.update(range(0, _WINDOW_HEAD))
        keep_ordinals.update(range(total_paras - _WINDOW_TAIL, total_paras))

    out = [
        "# Document outline (paragraph addresses) — WINDOWED (large document)",
        "# Showing every heading and table, plus a window of full paragraphs "
        "around the referenced location.",
        "# Ask to see a specific paragraph range (e.g. \"show paras 200-260\") "
        "for anything shown as omitted.",
    ]
    omitted = 0
    shown_paras = 0

    def _flush() -> None:
        nonlocal omitted
        if omitted:
            out.append(
                f"… ({omitted} paragraphs omitted; ask to see a specific range)"
            )
            omitted = 0

    for e in entries:
        if "para" not in e:
            _flush()
            out.append(_outline_line(e))  # tables / charts: always shown
            continue
        if _is_heading(e["style"]) or para_ordinal[e["para"]] in keep_ordinals:
            _flush()
            out.append(_outline_line(e))
            shown_paras += 1
        else:
            omitted += 1
    _flush()

    warnings.append(
        WarningEntry(
            code="outline_windowed",
            message=(
                f"Document outline ({total_paras} paragraphs) exceeded the "
                f"{_OUTLINE_CHAR_BUDGET}-char budget; sent a windowed view "
                f"({shown_paras} paragraphs shown). Ask to see a specific "
                "range for the rest."
            ),
            detail={
                "total_paragraphs": total_paras,
                "shown_paragraphs": shown_paras,
                "char_budget": _OUTLINE_CHAR_BUDGET,
                "anchored": bool(anchors),
            },
        )
    )
    return "\n".join(out)


_PARA_REF = re.compile(r"\bpara(?:graph)?\.?\s*#?\s*(\d+)", re.IGNORECASE)
_PARA_REF_KO = re.compile(r"(\d+)\s*번?\s*(?:문단|단락)")
_QUOTED = re.compile(r"[\"'“”‘’]([^\"'“”‘’]{4,})[\"'“”‘’]")


def _anchor_para_indices(req: EditDocRequest, entries: list[dict]) -> set[int]:
    """Paragraph indices the instruction/chat points at, to center a window.

    Two signals: an explicit paragraph number ("para 250", "250번 문단") and
    quoted text that matches a paragraph's body. Only indices that actually
    exist in the outline are returned.
    """
    text_parts = [req.instruction]
    text_parts += [t.content for t in req.chat_history[-12:]]
    text = "\n".join(text_parts)

    valid = {e["para"] for e in entries if "para" in e}
    anchors: set[int] = set()
    for m in _PARA_REF.finditer(text):
        n = int(m.group(1))
        if n in valid:
            anchors.add(n)
    for m in _PARA_REF_KO.finditer(text):
        n = int(m.group(1))
        if n in valid:
            anchors.add(n)

    quotes = [q.strip().lower() for q in _QUOTED.findall(text)]
    if quotes:
        for e in entries:
            if "para" not in e:
                continue
            body = e["text"].lower()
            if any(q and q in body for q in quotes):
                anchors.add(e["para"])
    return anchors


def _build_user_message(req: EditDocRequest, outline: str) -> str:
    lines = [outline, ""]
    if req.sources_markdown:
        lines.append("# Reference documents (attached to this turn)")
        for i, md in enumerate(req.sources_markdown, start=1):
            body = md.strip()
            if len(body) > 6000:
                body = body[:6000] + "\n…(truncated)"
            lines.append(f"## Document {i}")
            lines.append("```markdown")
            lines.append(body)
            lines.append("```")
        lines.append("")
    if req.chat_history:
        lines.append("# Chat history (most recent last)")
        for turn in req.chat_history[-12:]:
            lines.append(f"[{turn.role}] {turn.content.strip()[:500]}")
        lines.append("")
    lines.append("# Instruction (this turn)")
    lines.append(req.instruction.strip())
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Apply
# ---------------------------------------------------------------------------


_VALID_ACTIONS = {
    "docx": ("replace", "insert_after", "delete"),
    "xlsx": ("set_cell", "append_rows", "add_sheet"),
    "pptx": VALID_PPTX_ACTIONS,
}


def _as_int(v):
    """Coerce an address field to int — LLMs sometimes emit '3' or 3.0 for a
    slide/shape/row. Non-numeric values pass through unchanged so the engine
    flags them rather than a silent crash."""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return v


def _as_num(v):
    """Coerce a numeric field (size_pt / inches) to float. A non-numeric value
    (e.g. '2in') becomes None — that coordinate is simply left unchanged — so it
    never reaches EMU math as a string (which would build a giant throwaway
    string before failing)."""
    if v is None or isinstance(v, bool):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _as_bool(v):
    """Coerce a bold/italic flag to bool — accepts true/false/1/0/'true'."""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("true", "yes", "1", "on"):
            return True
        if s in ("false", "no", "0", "off"):
            return False
    return v


def _apply(
    fmt: str, content: bytes, raw_ops: list
) -> tuple[bytes, list[dict], list[WarningEntry], list[tuple[dict, str]]]:
    """Returns (bytes, applied_summaries, warnings, [(raw_op, status), ...]).

    The last element carries every VALID op paired with its result status,
    in input order, so the caller can stream per-op events with targets.
    """
    warnings: list[WarningEntry] = []
    valid_raw: list[dict] = []
    if fmt == "pptx":
        edits = []
        for raw in raw_ops:
            if not isinstance(raw, dict) or raw.get("action") not in VALID_PPTX_ACTIONS:
                warnings.append(_skip_warning(raw))
                continue
            valid_raw.append(raw)
            edits.append(
                PptxEdit(
                    action=raw["action"],
                    slide=_as_int(raw.get("slide")),
                    shape=_as_int(raw.get("shape")),
                    para=_as_int(raw.get("para")),  # None = whole shape / first para
                    row=_as_int(raw.get("row")),
                    col=_as_int(raw.get("col")),
                    at=_as_int(raw.get("at")),
                    chart=_as_int(raw.get("chart")),
                    new_text=str(raw["new_text"]) if raw.get("new_text") is not None else "",
                    old_text=raw.get("old_text"),
                    title=raw.get("title"),
                    categories=raw.get("categories"),
                    series=raw.get("series"),
                    color=raw.get("color"),
                    size_pt=_as_num(raw.get("size_pt")),
                    bold=_as_bool(raw.get("bold")),
                    italic=_as_bool(raw.get("italic")),
                    fill=raw.get("fill"),
                    left=_as_num(raw.get("left")),
                    top=_as_num(raw.get("top")),
                    width=_as_num(raw.get("width")),
                    height=_as_num(raw.get("height")),
                    runs=raw.get("runs"),
                    row2=_as_int(raw.get("row2")),
                    col2=_as_int(raw.get("col2")),
                    after=_as_int(raw.get("after")),
                    layout=raw.get("layout"),
                    body=(str(raw["body"]) if raw.get("body") is not None else None),
                    url=raw.get("url"),
                    bullet=raw.get("bullet"),
                    order=raw.get("order"),
                    run=_as_int(raw.get("run")),
                    position=raw.get("position"),
                    series_index=_as_int(raw.get("series_index")),
                    theme_name=raw.get("theme_name"),
                    typeface=raw.get("typeface"),
                    which=raw.get("which"),
                )
            )
        new_content, results = apply_pptx_edits(content, edits)
    elif fmt == "docx":
        edits = []
        for raw in raw_ops:
            if not isinstance(raw, dict) or raw.get("action") not in (
                "replace", "insert_after", "delete",
            ):
                warnings.append(_skip_warning(raw))
                continue
            valid_raw.append(raw)
            edits.append(
                DocxEdit(
                    action=raw["action"],
                    para=raw.get("para"),
                    table=raw.get("table"),
                    row=raw.get("row"),
                    col=raw.get("col"),
                    # `or ""` would erase falsy-but-real values like 0.
                    new_text=str(raw["new_text"]) if raw.get("new_text") is not None else "",
                    old_text=raw.get("old_text"),
                    markdown=str(raw["markdown"]) if raw.get("markdown") is not None else "",
                )
            )
        new_content, results = apply_docx_edits(content, edits)
    else:
        edits = []
        for raw in raw_ops:
            if not isinstance(raw, dict) or raw.get("action") not in (
                "set_cell", "append_rows", "add_sheet",
            ):
                warnings.append(_skip_warning(raw))
                continue
            valid_raw.append(raw)
            edits.append(
                XlsxEdit(
                    action=raw["action"],
                    sheet=str(raw["sheet"]) if raw.get("sheet") is not None else "",
                    cell=raw.get("cell"),
                    value=raw.get("value"),
                    old_value=raw.get("old_value"),
                    rows=raw.get("rows"),
                    headers=raw.get("headers"),
                )
            )
        new_content, results = apply_xlsx_edits(content, edits)

    applied: list[dict] = []
    op_results: list[tuple[dict, str]] = []
    # valid_raw and results are built in lockstep (one result per accepted
    # edit); strict=False preserves the pre-existing lenient pairing.
    for raw, result in zip(valid_raw, results, strict=False):
        op_results.append((raw, result.status))
        summary = {"action": result.action, "status": result.status}
        if result.status == "applied":
            applied.append(summary)
        else:
            warnings.append(
                WarningEntry(
                    code=f"edit_op_{result.status}",
                    message=f"{result.action} skipped ({result.status}): {result.message}",
                )
            )
    return (new_content if applied else content), applied, warnings, op_results


def _skip_warning(raw) -> WarningEntry:
    return WarningEntry(
        code="edit_op_unknown_action",
        message=f"Skipped op with unknown/invalid action: {raw!r}",
    )
