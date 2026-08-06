"""Extension-scoped tool exposure — tool_specs(extension=) + doc_guide(fmt=).

Scoping is opt-in: the unscoped defaults must stay byte-identical to today,
and a scoped view must drop only the verbs that don't apply and specialize
the descriptions of those that do.
"""

from __future__ import annotations

from edit2docs import tool_matrix
from edit2docs.agent_guide import doc_guide
from edit2docs.agent_tools import ANTHROPIC_TOOLS, TOOL_NAMES, tool_specs


class TestUnscopedUnchanged:
    def test_default_is_the_full_anthropic_set(self):
        assert tool_specs() is ANTHROPIC_TOOLS
        assert tool_specs("anthropic") is ANTHROPIC_TOOLS

    def test_openai_still_works_positionally(self):
        specs = tool_specs("openai")
        assert all(s["type"] == "function" for s in specs)
        assert [s["function"]["name"] for s in specs] == TOOL_NAMES

    def test_fmt_alias_back_compat(self):
        assert [s["function"]["name"] for s in tool_specs(fmt="openai")] == TOOL_NAMES


class TestVerbFiltering:
    def test_docx_drops_arrange_doc(self):
        names = [t["name"] for t in tool_specs(extension="docx")]
        assert "arrange_doc" not in names
        # everything else survives
        assert set(names) == set(TOOL_NAMES) - {"arrange_doc"}

    def test_xlsx_and_pptx_keep_arrange_doc(self):
        for ext in ("xlsx", "pptx", "deck.pptx", ".xlsx"):
            names = [t["name"] for t in tool_specs(extension=ext)]
            assert "arrange_doc" in names
            assert set(names) == set(TOOL_NAMES)

    def test_mixed_extensions_take_the_union(self):
        names = [t["name"] for t in tool_specs(extension={"docx", "pptx"})]
        # arrange_doc applies to pptx → present in the union
        assert set(names) == set(TOOL_NAMES)

    def test_unknown_extension_falls_back_to_full_set(self):
        names = [t["name"] for t in tool_specs(extension="txt")]
        assert set(names) == set(TOOL_NAMES)


class TestDescriptionSpecialization:
    def test_set_doc_text_scoped_to_format(self):
        xlsx = _desc("set_doc_text", "xlsx")
        assert "set_cell" in xlsx and "slide" not in xlsx and "chart" not in xlsx
        pptx = _desc("set_doc_text", "pptx")
        assert "chart" in pptx and "set_cell" not in pptx
        docx = _desc("set_doc_text", "docx")
        assert "insert_after" in docx and "set_cell" not in docx

    def test_arrange_doc_ops_scoped(self):
        assert "rename" in _desc("arrange_doc", "xlsx")
        assert "slide" in _desc("arrange_doc", "pptx")
        assert "rename" not in _desc("arrange_doc", "pptx")

    def test_build_doc_spec_scoped(self):
        assert "MARKDOWN" in _desc("build_doc", "docx")
        assert "sheets" in _desc("build_doc", "xlsx")
        assert "slides" in _desc("build_doc", "pptx")

    def test_mixed_extension_keeps_generic_description(self):
        got = _desc("set_doc_text", {"docx", "xlsx"})
        base = next(t for t in ANTHROPIC_TOOLS if t["name"] == "set_doc_text")
        assert got == base["description"]

    def test_scoped_descriptions_stay_compact(self):
        for ext in ("docx", "xlsx", "pptx"):
            for t in tool_specs(extension=ext):
                assert len(t["description"]) <= 320, (t["name"], ext)

    def test_openai_shape_scoped(self):
        specs = tool_specs("openai", extension="xlsx")
        d = next(s for s in specs if s["function"]["name"] == "set_doc_text")
        assert "set_cell" in d["function"]["description"]


class TestDocGuideScoping:
    def test_unscoped_topics_unchanged(self):
        assert doc_guide()["topics"] == doc_guide(fmt=None)["topics"]

    def test_docx_topics_drop_slide_and_arrange(self):
        topics = doc_guide(fmt="docx")["topics"]
        assert "arrange" not in topics
        assert "recipes.slides" not in topics
        assert "edit.chart" not in topics  # charts are xlsx/pptx
        assert "build" in topics and "edit.text" in topics

    def test_pptx_keeps_slide_topics(self):
        topics = doc_guide(fmt="pptx")["topics"]
        assert {"arrange", "recipes.slides", "edit.chart"} <= set(topics)

    def test_xlsx_keeps_arrange_not_slides(self):
        topics = doc_guide(fmt="xlsx")["topics"]
        assert "arrange" in topics and "edit.chart" in topics
        assert "recipes.slides" not in topics

    def test_topic_guide_text_still_served_when_scoped(self):
        g = doc_guide("arrange", fmt="pptx")
        assert "arrange_doc" in g["guide"]


class TestMatrixHelpers:
    def test_applicable_verbs(self):
        assert "arrange_doc" not in tool_matrix.applicable_verbs("docx")
        assert tool_matrix.applicable_verbs(None) is None

    def test_describe(self):
        assert "set_cell" in tool_matrix.describe("set_doc_text", "xlsx")
        assert tool_matrix.describe("set_doc_text", {"docx", "xlsx"}) is None
        assert tool_matrix.describe("analyze_doc", "xlsx") is None  # no override


def _desc(verb: str, extension) -> str:
    return next(t for t in tool_specs(extension=extension) if t["name"] == verb)[
        "description"
    ]
