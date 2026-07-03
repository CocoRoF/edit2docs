"""Chat-edit orchestrator for DOCX / XLSX (one turn = one call).

Same contract as the PPTX chat editor (tools/edit_deck.py): a planner LLM
turns the instruction + a structural outline into a minimal operation
list plus a chat reply; the deterministic engines apply the operations so
untouched content survives byte-identical. Question-only turns answer
without changing the file. Plan-block failures retry once and then admit
failure in the reply instead of promising changes.
"""

from __future__ import annotations

import time
from typing import Literal

from pydantic import Field

from ..documents.docx_engine import DocxEdit, apply_docx_edits, docx_outline
from ..documents.xlsx_engine import XlsxEdit, apply_xlsx_edits, xlsx_outline
from ..llm import AnthropicClient, DEFAULT_MODEL, build_output_lang_directive, load_prompt
from ._edit_events import op_event_vars, op_summary, plan_event_vars
from .edit_deck import ChatTurn, _cost_from_usage, _parse_plan
from .generate_deck import EventCallback, StageEvent, _emit, _merge_cost
from .types import (
    CostBreakdown,
    DEFAULT_LANG,
    LangCode,
    ToolRequest,
    ToolResponse,
    WarningEntry,
)

DocFormat = Literal["docx", "xlsx"]

_PLANNER_ROLE = {"docx": "doc-editor-planner", "xlsx": "sheet-editor-planner"}
_MAX_OPERATIONS = 30


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
    outline_text = _outline_context(req)
    client = AnthropicClient(api_key=req.anthropic_api_key, model=req.model)
    system = (
        build_output_lang_directive(req.lang)
        + "\n\n"
        + load_prompt(_PLANNER_ROLE[req.fmt])
    )
    user = _build_user_message(req, outline_text)

    result = await client.complete(
        system_prompt=system,
        user_message=user,
        max_output_tokens=16384,
        cache_system=True,
        model=req.model,
    )
    cost = _merge_cost(cost, _cost_from_usage(result.usage))
    reply, raw_ops, plan_missing = _parse_plan(result.text, warnings)

    if plan_missing:
        retry = await client.complete(
            system_prompt=system,
            user_message=(
                user
                + "\n\n# REMINDER\nYour previous answer was missing the "
                "```edit_plan fenced block. Respond again following the "
                "output format EXACTLY (operations: [] only if truly no "
                "change is needed)."
            ),
            max_output_tokens=16384,
            cache_system=True,
            model=req.model,
        )
        cost = _merge_cost(cost, _cost_from_usage(retry.usage))
        reply, raw_ops, plan_missing = _parse_plan(retry.text, warnings)
        if plan_missing:
            reply = (
                reply.rstrip()
                + "\n\n[주의] 편집 계획 생성에 실패해 변경이 적용되지 않았습니다. "
                "요청을 조금 더 구체적으로 나눠서 다시 보내주세요."
            )

    if len(raw_ops) > _MAX_OPERATIONS:
        warnings.append(
            WarningEntry(
                code="edit_plan_truncated",
                message=f"{len(raw_ops)} operations planned; applying first {_MAX_OPERATIONS}.",
                detail={"emitted": len(raw_ops), "cap": _MAX_OPERATIONS},
            )
        )
        reply = (
            reply.rstrip()
            + f"\n\n[안내] 계획된 작업 {len(raw_ops)}개 중 상한에 따라 앞 "
            f"{_MAX_OPERATIONS}개만 이번 턴에 적용합니다. 같은 요청을 한 번 더 "
            "보내면 이어서 처리됩니다."
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
                message_vars=plan_event_vars(req.fmt, valid_ops),
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
                        op_summary(req.fmt, op, index=i, total=total),
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


def _outline_context(req: EditDocRequest) -> str:
    if req.fmt == "docx":
        lines = ["# Document outline (paragraph addresses)"]
        for entry in docx_outline(req.content):
            if "para" in entry:
                lines.append(
                    f"- para {entry['para']} [{entry['style']}]: {entry['text'][:160]}"
                )
            else:
                lines.append(
                    f"- table {entry['table']} cell ({entry['row']},{entry['col']}): "
                    f"{entry['text'][:120]}"
                )
        return "\n".join(lines)

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
}


def _apply(
    fmt: str, content: bytes, raw_ops: list
) -> tuple[bytes, list[dict], list[WarningEntry], list[tuple[dict, str]]]:
    """Returns (bytes, applied_summaries, warnings, [(raw_op, status), ...]).

    The last element carries every VALID op paired with its result status,
    in input order, so the caller can stream per-op events with targets.
    """
    warnings: list[WarningEntry] = []
    valid_raw: list[dict] = []
    if fmt == "docx":
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
    for raw, result in zip(valid_raw, results):
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
