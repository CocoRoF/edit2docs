"""Operation-level streaming for the chat editors (live edit view).

The web studio subscribes to the job's SSE stream. Beyond coarse stage
labels, it wants to show *which part* is being edited and the edit itself
as it happens. We carry that in ``StageEvent.message_vars`` (already
streamed end-to-end) under two keys:

* ``plan`` — emitted once after planning: a list of ``op_summary`` dicts so
  the UI can show the whole todo list up front.
* ``op``  — emitted per operation as it starts and finishes:
  ``{**op_summary, "phase": "start"|"done", "status": ...}``.

``op_summary`` is format-agnostic on the wire::

    {
      "index": 0, "total": 3,
      "action": "replace",
      "target": {"kind": "paragraph", "para": 2},   # address the UI can locate
      "label": "2번 문단 교체",                        # ko human label
    }

Keeping this in message_vars means no schema/model change and no new event
type — old clients simply ignore the extra keys.
"""

from __future__ import annotations

from typing import Any

__all__ = ["op_summary", "plan_event_vars", "op_event_vars"]


def _pptx_target(op: dict) -> tuple[dict, str]:
    action = op.get("action")
    if action == "edit":
        n = op.get("slide")
        return {"kind": "slide", "slide": n}, f"{n}번 슬라이드 편집"
    if action == "add":
        after = op.get("after", 0)
        return {"kind": "slide_after", "after": after}, (
            "맨 앞에 새 슬라이드 추가" if after == 0 else f"{after}번 뒤 새 슬라이드 추가"
        )
    if action == "delete":
        n = op.get("slide")
        return {"kind": "slide", "slide": n}, f"{n}번 슬라이드 삭제"
    return {"kind": "unknown"}, action or "작업"


def _docx_target(op: dict) -> tuple[dict, str]:
    action = op.get("action")
    if op.get("table") is not None:
        t, r, c = op.get("table"), op.get("row"), op.get("col")
        return (
            {"kind": "table_cell", "table": t, "row": r, "col": c},
            f"표 {t} 셀({r},{c}) 교체",
        )
    para = op.get("para")
    if action == "replace":
        return {"kind": "paragraph", "para": para}, f"{para}번 문단 교체"
    if action == "insert_after":
        return {"kind": "paragraph_after", "para": para}, (
            "맨 앞에 내용 삽입" if para == -1 else f"{para}번 문단 뒤 삽입"
        )
    if action == "delete":
        return {"kind": "paragraph", "para": para}, f"{para}번 문단 삭제"
    return {"kind": "unknown"}, action or "작업"


def _xlsx_target(op: dict) -> tuple[dict, str]:
    action = op.get("action")
    sheet = op.get("sheet", "")
    if action == "set_cell":
        cell = op.get("cell", "")
        return {"kind": "cell", "sheet": sheet, "cell": cell}, f"[{sheet}] {cell} 셀 수정"
    if action == "append_rows":
        n = len(op.get("rows") or [])
        return {"kind": "sheet", "sheet": sheet}, f"[{sheet}] {n}개 행 추가"
    if action == "add_sheet":
        return {"kind": "sheet", "sheet": sheet}, f"[{sheet}] 시트 추가"
    return {"kind": "sheet", "sheet": sheet}, action or "작업"


_TARGETERS = {"pptx": _pptx_target, "docx": _docx_target, "xlsx": _xlsx_target}


def op_summary(fmt: str, op: dict, *, index: int, total: int) -> dict[str, Any]:
    """Format-agnostic descriptor for one planned/applied operation."""
    target, label = _TARGETERS.get(fmt, _pptx_target)(op)
    return {
        "index": index,
        "total": total,
        "action": op.get("action"),
        "target": target,
        "label": label,
    }


def plan_event_vars(fmt: str, ops: list[dict]) -> dict[str, Any]:
    """message_vars payload announcing the full edit plan."""
    total = len(ops)
    return {
        "plan": [op_summary(fmt, op, index=i, total=total) for i, op in enumerate(ops)]
    }


def op_event_vars(
    summary: dict, *, phase: str, status: str | None = None
) -> dict[str, Any]:
    """message_vars payload for one operation's start/done."""
    op = {**summary, "phase": phase}
    if status is not None:
        op["status"] = status
    return {"op": op}
