"""End-to-end tests for the chat-edit pipeline (tools.edit_deck).

LLM calls are stubbed (routing on the system prompt: planner vs slide
editor); everything deterministic — preview render, plan validation,
recompose — runs for real against a python-pptx-built host deck.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu, Inches

from edit2docs.llm.anthropic_client import LLMResult, LLMUsage
from edit2docs.tools.edit_deck import EditDeckRequest, edit_deck

NEW_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">'
    '<rect x="0" y="0" width="1280" height="720" fill="#F7F9FC"/>'
    '<text x="120" y="170" font-size="40" fill="#111111">채팅으로 편집된 슬라이드</text>'
    "</svg>"
)

PLAN_EDIT_AND_ADD = (
    "```reply\n2번 슬라이드를 수정하고 마지막에 새 슬라이드를 추가합니다.\n```\n"
    "```edit_plan\n"
    "operations:\n"
    '  - action: edit\n    slide: 2\n    brief: "제목 교체"\n'
    '  - action: add\n    after: 3\n    brief: "요약 슬라이드"\n'
    '  - action: delete\n    slide: 1\n'
    "```"
)

PLAN_EMPTY = (
    "```reply\n이 덱은 3장이고 표지·본문·결론으로 구성되어 있습니다.\n```\n"
    "```edit_plan\noperations: []\n```"
)


@dataclass
class _RoutingLLM:
    """Planner calls get `plan`; slide-editor calls get an SVG block."""

    plan: str
    calls: list[dict] = field(default_factory=list)

    async def complete(self, system_prompt, user_message, **kwargs):
        self.calls.append({"system": system_prompt, "user": user_message})
        if "Deck Edit Planner" in system_prompt:
            text = self.plan
        else:
            text = f"```svg\n{NEW_SVG}\n```"
        return LLMResult(
            text=text,
            usage=LLMUsage(input_tokens=10, output_tokens=10),
            model="stub",
            stop_reason="end_turn",
        )


def _host_pptx_bytes(tmp_path: Path) -> bytes:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    for title in ("표지", "본문", "결론"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        box.text_frame.text = title
    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path.read_bytes()


def _texts(pptx: bytes, tmp_path: Path) -> list[str]:
    out = tmp_path / "check.pptx"
    out.write_bytes(pptx)
    prs = Presentation(str(out))
    return [
        "".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
        for s in prs.slides
    ]


def _request(pptx: bytes, instruction: str = "2번 슬라이드 고쳐줘") -> EditDeckRequest:
    return EditDeckRequest(
        pptx=pptx,
        instruction=instruction,
        lang="ko-KR",
        anthropic_api_key="sk-ant-stub",
    )


class TestEditDeck:
    @pytest.mark.asyncio
    async def test_edit_add_delete_turn(self, monkeypatch, tmp_path):
        llm = _RoutingLLM(plan=PLAN_EDIT_AND_ADD)
        import sys

        ed = sys.modules["edit2docs.tools.edit_deck"]
        monkeypatch.setattr(ed, "AnthropicClient", lambda **kw: llm)

        events: list[str] = []
        resp = await edit_deck(
            _request(_host_pptx_bytes(tmp_path)),
            on_event=lambda e: events.append(e.stage),
        )

        assert resp.changed is True
        # delete slide1, edit slide2, keep slide3, add one at the end -> 3 slides
        texts = _texts(resp.pptx, tmp_path)
        assert len(texts) == 3
        assert "채팅으로 편집된" in texts[0]  # edited slide 2 is now first
        assert texts[1] == "결론"
        assert "채팅으로 편집된" in texts[2]  # appended slide

        assert resp.reply.startswith("2번 슬라이드")
        assert {op["action"] for op in resp.operations} == {"edit", "add", "delete"}
        for stage in ("analyzing_deck", "planning_edits", "editing_slides", "applying_edits", "done"):
            assert stage in events, events

        # Planner saw the outline; slide editor got the stubbed SVG task.
        assert "# Deck outline" in llm.calls[0]["user"]
        assert len(llm.calls) == 3  # 1 planner + 2 slide generations

    @pytest.mark.asyncio
    async def test_question_only_turn_leaves_deck_untouched(self, monkeypatch, tmp_path):
        llm = _RoutingLLM(plan=PLAN_EMPTY)
        import sys

        ed = sys.modules["edit2docs.tools.edit_deck"]
        monkeypatch.setattr(ed, "AnthropicClient", lambda **kw: llm)

        original = _host_pptx_bytes(tmp_path)
        resp = await edit_deck(_request(original, "이 덱 구성이 어떻게 돼?"))

        assert resp.changed is False
        assert resp.pptx == original
        assert resp.operations == []
        assert "3장" in resp.reply
        assert len(llm.calls) == 1  # planner only

    @pytest.mark.asyncio
    async def test_chat_history_reaches_planner(self, monkeypatch, tmp_path):
        llm = _RoutingLLM(plan=PLAN_EMPTY)
        import sys

        ed = sys.modules["edit2docs.tools.edit_deck"]
        monkeypatch.setattr(ed, "AnthropicClient", lambda **kw: llm)

        from edit2docs.tools.edit_deck import ChatTurn

        req = EditDeckRequest(
            pptx=_host_pptx_bytes(tmp_path),
            instruction="아까 말한 색으로 바꿔줘",
            chat_history=[
                ChatTurn(role="user", content="파란색 계열이 좋아"),
                ChatTurn(role="assistant", content="네, 파란색 팔레트로 기억할게요."),
            ],
            lang="ko-KR",
            anthropic_api_key="sk-ant-stub",
        )
        await edit_deck(req)
        planner_user = llm.calls[0]["user"]
        assert "파란색 계열" in planner_user
        assert "# Chat history" in planner_user
