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
      "label": "Replace paragraph 2",               # human label in the turn's lang
    }

Labels are localized to the edit turn's ``lang`` (English-first; Korean
fully supported). Keeping this in message_vars means no schema/model change
and no new event type — old clients simply ignore the extra keys.
"""

from __future__ import annotations

from typing import Any

__all__ = ["op_summary", "plan_event_vars", "op_event_vars"]


def _is_ko(lang: str) -> bool:
    return (lang or "").lower().startswith("ko")


def _pptx_target(op: dict, lang: str) -> tuple[dict, str]:
    action = op.get("action")
    ko = _is_ko(lang)
    n = op.get("slide")

    # -- surgical structural ops (documents/pptx_engine) --------------------
    if action == "set_text":
        shape = op.get("shape")
        label = f"{n}번 슬라이드 텍스트 편집" if ko else f"Edit text on slide {n}"
        return {"kind": "shape", "slide": n, "shape": shape}, label
    if action == "set_table_cell" or action == "set_cell_style":
        shape, r, c = op.get("shape"), op.get("row"), op.get("col")
        if action == "set_cell_style":
            label = (
                f"{n}번 슬라이드 표 셀({r},{c}) 서식" if ko
                else f"Style table cell ({r},{c}) on slide {n}"
            )
        else:
            label = (
                f"{n}번 슬라이드 표 셀({r},{c}) 편집" if ko
                else f"Edit table cell ({r},{c}) on slide {n}"
            )
        return {"kind": "table_cell", "slide": n, "shape": shape, "row": r, "col": c}, label
    if action == "set_shape_style":
        shape = op.get("shape")
        label = f"{n}번 슬라이드 서식 변경" if ko else f"Restyle a shape on slide {n}"
        return {"kind": "shape", "slide": n, "shape": shape}, label
    if action == "set_shape_position":
        shape = op.get("shape")
        label = f"{n}번 슬라이드 요소 이동" if ko else f"Move a shape on slide {n}"
        return {"kind": "shape", "slide": n, "shape": shape}, label
    if action == "insert_row":
        shape = op.get("shape")
        label = f"{n}번 슬라이드 표에 행 추가" if ko else f"Add a table row on slide {n}"
        return {"kind": "table", "slide": n, "shape": shape}, label
    if action == "delete_row":
        shape, r = op.get("shape"), op.get("row")
        label = f"{n}번 슬라이드 표 {r}행 삭제" if ko else f"Delete table row {r} on slide {n}"
        return {"kind": "table", "slide": n, "shape": shape}, label
    if action == "set_chart_data":
        chart = op.get("chart")
        label = f"{n}번 슬라이드 차트 데이터 수정" if ko else f"Update chart data on slide {n}"
        return {"kind": "chart", "slide": n, "chart": chart}, label
    if action == "set_chart_title":
        chart = op.get("chart")
        label = f"{n}번 슬라이드 차트 제목 수정" if ko else f"Update chart title on slide {n}"
        return {"kind": "chart", "slide": n, "chart": chart}, label

    # -- slide-level ops (SVG editor + redesign fallback) ------------------
    if action == "edit":
        label = f"{n}번 슬라이드 편집" if ko else f"Edit slide {n}"
        return {"kind": "slide", "slide": n}, label
    if action == "redesign":
        label = f"{n}번 슬라이드 재구성" if ko else f"Redesign slide {n}"
        return {"kind": "slide", "slide": n}, label
    if action == "add":
        after = op.get("after", 0)
        if after == 0:
            label = "맨 앞에 새 슬라이드 추가" if ko else "Add a new slide at the start"
        else:
            label = f"{after}번 뒤 새 슬라이드 추가" if ko else f"Add a new slide after {after}"
        return {"kind": "slide_after", "after": after}, label
    if action == "delete":
        label = f"{n}번 슬라이드 삭제" if ko else f"Delete slide {n}"
        return {"kind": "slide", "slide": n}, label
    return {"kind": "unknown"}, action or ("작업" if ko else "operation")


def _docx_target(op: dict, lang: str) -> tuple[dict, str]:
    action = op.get("action")
    ko = _is_ko(lang)
    if op.get("table") is not None:
        t, r, c = op.get("table"), op.get("row"), op.get("col")
        label = (
            f"표 {t} 셀({r},{c}) 교체" if ko else f"Replace table {t} cell ({r},{c})"
        )
        return {"kind": "table_cell", "table": t, "row": r, "col": c}, label
    para = op.get("para")
    if action == "replace":
        label = f"{para}번 문단 교체" if ko else f"Replace paragraph {para}"
        return {"kind": "paragraph", "para": para}, label
    if action == "insert_after":
        if para == -1:
            label = "맨 앞에 내용 삽입" if ko else "Insert content at the start"
        else:
            label = f"{para}번 문단 뒤 삽입" if ko else f"Insert after paragraph {para}"
        return {"kind": "paragraph_after", "para": para}, label
    if action == "delete":
        label = f"{para}번 문단 삭제" if ko else f"Delete paragraph {para}"
        return {"kind": "paragraph", "para": para}, label
    return {"kind": "unknown"}, action or ("작업" if ko else "operation")


def _xlsx_target(op: dict, lang: str) -> tuple[dict, str]:
    action = op.get("action")
    sheet = op.get("sheet", "")
    ko = _is_ko(lang)
    if action == "set_cell":
        cell = op.get("cell", "")
        label = f"[{sheet}] {cell} 셀 수정" if ko else f"[{sheet}] Edit cell {cell}"
        return {"kind": "cell", "sheet": sheet, "cell": cell}, label
    if action == "append_rows":
        n = len(op.get("rows") or [])
        label = f"[{sheet}] {n}개 행 추가" if ko else f"[{sheet}] Append {n} row(s)"
        return {"kind": "sheet", "sheet": sheet}, label
    if action == "add_sheet":
        label = f"[{sheet}] 시트 추가" if ko else f"Add sheet [{sheet}]"
        return {"kind": "sheet", "sheet": sheet}, label
    return {"kind": "sheet", "sheet": sheet}, action or ("작업" if ko else "operation")


_TARGETERS = {"pptx": _pptx_target, "docx": _docx_target, "xlsx": _xlsx_target}


def op_summary(
    fmt: str, op: dict, *, index: int, total: int, lang: str = "en-US"
) -> dict[str, Any]:
    """Format-agnostic descriptor for one planned/applied operation."""
    target, label = _TARGETERS.get(fmt, _pptx_target)(op, lang)
    return {
        "index": index,
        "total": total,
        "action": op.get("action"),
        "target": target,
        "label": label,
    }


def plan_event_vars(fmt: str, ops: list[dict], *, lang: str = "en-US") -> dict[str, Any]:
    """message_vars payload announcing the full edit plan."""
    total = len(ops)
    return {
        "plan": [
            op_summary(fmt, op, index=i, total=total, lang=lang)
            for i, op in enumerate(ops)
        ]
    }


def op_event_vars(
    summary: dict, *, phase: str, status: str | None = None
) -> dict[str, Any]:
    """message_vars payload for one operation's start/done."""
    op = {**summary, "phase": phase}
    if status is not None:
        op["status"] = status
    return {"op": op}
