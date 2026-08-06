"""Structural PPTX editor (documents/pptx_engine): addressable outline +
surgical, byte-preserving operations on real decks (text / table / chart).

These are the guarantees the legacy SVG-rewrite path could not give: tables
and charts are edited IN PLACE (not flattened), and everything an edit does
not touch stays byte-identical.
"""

from __future__ import annotations

import io
import zipfile

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Inches, Pt

from edit2docs.documents.pptx_engine import (
    PptxEdit,
    apply_pptx_edits,
    pptx_outline,
)


def _deck() -> bytes:
    """Slide 1: title textbox + 3x3 table + column chart. Slide 2: text only."""
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s1.shapes.add_textbox(Inches(1), Inches(0.3), Inches(6), Inches(1))
    tb.text_frame.text = "원본 제목"
    tbl = s1.shapes.add_table(3, 3, Inches(0.5), Inches(1.5), Inches(6), Inches(2)).table
    for r in range(3):
        for c in range(3):
            tbl.cell(r, c).text = f"R{r}C{c}"
    cd = CategoryChartData()
    cd.categories = ["Q1", "Q2", "Q3"]
    cd.add_series("매출", (10.0, 20.0, 30.0))
    s1.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(3.8), Inches(5), Inches(3), cd
    )

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    b2 = s2.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    b2.text_frame.text = "둘째 슬라이드"

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _ids(deck: bytes):
    outline = pptx_outline(deck)
    title = next(e for e in outline if e.get("kind") == "text" and e["slide"] == 1)
    table = next(e for e in outline if e.get("kind") == "table")
    chart = next(e for e in outline if e.get("kind") == "chart")
    return title["shape"], table["shape"], chart["chart"]


def _members(pptx: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(pptx)) as z:
        return {n: z.read(n) for n in z.namelist()}


class TestOutline:
    def test_outline_lists_text_table_and_chart(self):
        outline = pptx_outline(_deck())
        kinds = {e.get("kind") for e in outline}
        assert {"text", "table", "chart"} <= kinds
        # table cells are addressable
        cells = [e for e in outline if "table" in e and "row" in e]
        assert len(cells) == 9
        assert {(e["row"], e["col"]) for e in cells} == {(r, c) for r in range(3) for c in range(3)}
        # chart series data is surfaced
        chart = next(e for e in outline if e.get("kind") == "chart")
        assert chart["series"][0]["categories"] == ["Q1", "Q2", "Q3"]
        assert chart["series"][0]["values"] == [10.0, 20.0, 30.0]


class TestSurgicalEdits:
    def test_all_op_types_apply(self):
        deck = _deck()
        title, table, _ = _ids(deck)
        edits = [
            PptxEdit("set_text", slide=1, shape=title, para=0, new_text="새 제목", old_text="원본 제목"),
            PptxEdit("set_table_cell", slide=1, shape=table, row=0, col=0, new_text="바뀐셀", old_text="R0C0"),
            PptxEdit("insert_row", slide=1, shape=table, at=3),
            PptxEdit("set_table_cell", slide=1, shape=table, row=3, col=0, new_text="새행"),
            PptxEdit("set_chart_data", slide=1, chart=0,
                     categories=["Q1", "Q2", "Q3", "Q4"], series=[{"name": "매출", "values": [11, 22, 33, 44]}]),
            PptxEdit("set_chart_title", slide=1, chart=0, title="분기 매출"),
        ]
        out, results = apply_pptx_edits(deck, edits)
        assert [r.status for r in results] == ["applied"] * 6

        prs = Presentation(io.BytesIO(out))
        s1 = prs.slides[0]
        texts = [sh.text_frame.text for sh in s1.shapes if sh.has_text_frame]
        assert "새 제목" in texts
        tbl = next(sh.table for sh in s1.shapes if sh.has_table)
        assert len(tbl.rows) == 4
        assert tbl.cell(0, 0).text == "바뀐셀"
        assert tbl.cell(3, 0).text == "새행"
        chart = next(sh.chart for sh in s1.shapes if sh.has_chart)
        assert chart.has_title and chart.chart_title.text_frame.text == "분기 매출"
        assert list(chart.plots[0].categories) == ["Q1", "Q2", "Q3", "Q4"]
        assert list(chart.series[0].values) == [11.0, 22.0, 33.0, 44.0]

    def test_untouched_parts_are_byte_identical(self):
        """Editing slide 1 must not rewrite slide 2 (or any unrelated part)."""
        deck = _deck()
        title, _, _ = _ids(deck)
        before = _members(deck)
        out, results = apply_pptx_edits(
            deck, [PptxEdit("set_text", slide=1, shape=title, para=0, new_text="X")]
        )
        assert results[0].status == "applied"
        after = _members(out)
        # slide2.xml is untouched → byte-identical
        assert after["ppt/slides/slide2.xml"] == before["ppt/slides/slide2.xml"]
        # the theme / presentation parts are untouched too
        assert after["ppt/theme/theme1.xml"] == before["ppt/theme/theme1.xml"]
        # slide1 DID change
        assert after["ppt/slides/slide1.xml"] != before["ppt/slides/slide1.xml"]

    def test_delete_row(self):
        deck = _deck()
        _, table, _ = _ids(deck)
        out, results = apply_pptx_edits(deck, [PptxEdit("delete_row", slide=1, shape=table, row=1)])
        assert results[0].status == "applied"
        tbl = next(sh.table for sh in Presentation(io.BytesIO(out)).slides[0].shapes if sh.has_table)
        assert len(tbl.rows) == 2
        assert tbl.cell(1, 0).text == "R2C0"  # row 1 gone, old row 2 shifted up


class TestGoldenDeckPreservation:
    """The core guarantee the SVG rewrite could not give: editing some content
    NEVER drops or alters the styled header, theme, or native objects the edit
    did not touch. This is the regression for '스타일 헤더가 맘대로 지워짐'."""

    def _styled_deck(self) -> tuple[bytes, int, int, int]:
        prs = Presentation()
        prs.slide_width = Emu(12192000)
        prs.slide_height = Emu(6858000)
        s = prs.slides.add_slide(prs.slide_layouts[6])
        # A styled "header" band we will NOT touch.
        hdr = s.shapes.add_textbox(Inches(0), Inches(0), Inches(13.3), Inches(0.8))
        hdr.text_frame.text = "COMPANY HEADER"
        run = hdr.text_frame.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(20)
        from pptx.dml.color import RGBColor
        run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        # A title we WILL edit.
        title = s.shapes.add_textbox(Inches(1), Inches(1.2), Inches(6), Inches(1))
        title.text_frame.text = "원본 제목"
        # A table we WILL edit one cell of.
        tbl = s.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(5), Inches(1.5)).table
        for r in range(2):
            for c in range(2):
                tbl.cell(r, c).text = f"R{r}C{c}"
        buf = io.BytesIO()
        prs.save(buf)
        deck = buf.getvalue()
        return deck, hdr.shape_id, title.shape_id, tbl_shape_id(deck)

    def test_untouched_styled_header_survives_edits_to_other_shapes(self):
        deck, hdr_id, title_id, table_id = self._styled_deck()
        out, results = apply_pptx_edits(deck, [
            PptxEdit("set_text", slide=1, shape=title_id, para=0, new_text="새 제목"),
            PptxEdit("set_table_cell", slide=1, shape=table_id, row=0, col=0, new_text="변경"),
        ])
        assert [r.status for r in results] == ["applied", "applied"]

        prs = Presentation(io.BytesIO(out))
        shapes = {sh.shape_id: sh for sh in prs.slides[0].shapes}
        # The header is byte-perfectly intact: text, bold, size, color.
        hdr = shapes[hdr_id]
        hrun = hdr.text_frame.paragraphs[0].runs[0]
        assert hdr.text_frame.text == "COMPANY HEADER"
        assert hrun.font.bold is True
        assert hrun.font.size == Pt(20)
        assert str(hrun.font.color.rgb) == "C00000"
        # The targeted edits landed.
        assert shapes[title_id].text_frame.text == "새 제목"
        tbl = next(sh.table for sh in prs.slides[0].shapes if sh.has_table)
        assert tbl.cell(0, 0).text == "변경" and tbl.cell(1, 1).text == "R1C1"


def tbl_shape_id(deck: bytes) -> int:
    return next(e["shape"] for e in pptx_outline(deck) if e.get("kind") == "table")


class TestStyleAndPosition:
    def test_outline_reports_shape_geometry_in_inches(self):
        outline = pptx_outline(_deck())
        title = next(e for e in outline if e.get("kind") == "text" and e["slide"] == 1)
        assert title["pos"] == {"left": 1.0, "top": 0.3, "width": 6.0, "height": 1.0}

    def test_set_shape_style(self):
        deck = _deck()
        title, _, _ = _ids(deck)
        out, results = apply_pptx_edits(deck, [
            PptxEdit("set_shape_style", slide=1, shape=title,
                     color="FF0000", size_pt=28, bold=True, fill="EEEEEE"),
        ])
        assert results[0].status == "applied"
        sh = next(s for s in Presentation(io.BytesIO(out)).slides[0].shapes if s.has_text_frame)
        run = sh.text_frame.paragraphs[0].runs[0]
        assert str(run.font.color.rgb) == "FF0000"
        assert run.font.size == 355600  # 28pt
        assert run.font.bold is True
        assert str(sh.fill.fore_color.rgb) == "EEEEEE"
        # text is unchanged
        assert sh.text_frame.text == "원본 제목"

    def test_set_shape_position(self):
        deck = _deck()
        title, _, _ = _ids(deck)
        out, results = apply_pptx_edits(deck, [
            PptxEdit("set_shape_position", slide=1, shape=title,
                     left=2.0, top=3.0, width=6.0, height=1.5),
        ])
        assert results[0].status == "applied"
        sh = next(s for s in Presentation(io.BytesIO(out)).slides[0].shapes if s.has_text_frame)
        assert (round(sh.left / 914400, 2), round(sh.top / 914400, 2),
                round(sh.width / 914400, 2), round(sh.height / 914400, 2)) == (2.0, 3.0, 6.0, 1.5)

    def test_set_cell_style(self):
        deck = _deck()
        _, table, _ = _ids(deck)
        out, results = apply_pptx_edits(deck, [
            PptxEdit("set_cell_style", slide=1, shape=table, row=0, col=0,
                     fill="0000FF", color="FFFFFF", bold=True),
        ])
        assert results[0].status == "applied"
        cell = next(
            sh.table for sh in Presentation(io.BytesIO(out)).slides[0].shapes if sh.has_table
        ).cell(0, 0)
        assert cell.text == "R0C0"  # text unchanged
        assert str(cell.fill.fore_color.rgb) == "0000FF"
        assert str(cell.text_frame.paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"

    def test_style_needs_an_attribute(self):
        deck = _deck()
        title, _, _ = _ids(deck)
        _, results = apply_pptx_edits(deck, [PptxEdit("set_shape_style", slide=1, shape=title)])
        assert results[0].status == "invalid"


class TestGuardsAndErrors:
    def test_stale_old_text_is_rejected(self):
        deck = _deck()
        title, table, _ = _ids(deck)
        _, results = apply_pptx_edits(deck, [
            PptxEdit("set_text", slide=1, shape=title, para=0, new_text="X", old_text="틀린 원본"),
            PptxEdit("set_table_cell", slide=1, shape=table, row=0, col=0, new_text="X", old_text="wrong"),
        ])
        assert results[0].status == "stale"
        assert results[1].status == "stale"

    def test_not_found_addresses(self):
        deck = _deck()
        _, results = apply_pptx_edits(deck, [
            PptxEdit("set_text", slide=9, shape=2, new_text="X"),        # bad slide
            PptxEdit("set_text", slide=1, shape=9999, new_text="X"),     # bad shape
            PptxEdit("set_chart_data", slide=1, chart=5, categories=["a"], series=[{"name": "s", "values": [1]}]),
        ])
        assert results[0].status == "not_found"
        assert results[1].status == "not_found"
        assert results[2].status == "not_found"

    def test_ragged_chart_data_is_invalid(self):
        deck = _deck()
        _, results = apply_pptx_edits(deck, [
            PptxEdit("set_chart_data", slide=1, chart=0,
                     categories=["Q1", "Q2"], series=[{"name": "s", "values": [1, 2, 3]}]),
        ])
        assert results[0].status == "invalid"

    def test_unknown_action_is_invalid(self):
        deck = _deck()
        _, results = apply_pptx_edits(deck, [PptxEdit("frobnicate", slide=1, shape=2)])
        assert results[0].status == "invalid"

    def test_batch_is_order_preserving_in_results(self):
        deck = _deck()
        title, table, _ = _ids(deck)
        _, results = apply_pptx_edits(deck, [
            PptxEdit("delete_row", slide=1, shape=table, row=2),
            PptxEdit("set_text", slide=1, shape=title, para=0, new_text="Z"),
        ])
        # results correlate to input order even though row-op sorts internally
        assert results[0].action == "delete_row" and results[0].status == "applied"
        assert results[1].action == "set_text" and results[1].status == "applied"
