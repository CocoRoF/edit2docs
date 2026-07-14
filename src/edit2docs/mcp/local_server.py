"""Zero-infra MCP server: edit2docs tools over LOCAL files, stdio transport.

Needs nothing but the pip package — tools read and write .docx/.xlsx/.pptx
files on the local filesystem directly:

    { "mcpServers": { "edit2docs": {
        "command": "edit2docs-mcp",
        "env": {"ANTHROPIC_API_KEY": "sk-ant-..."}
    } } }

``generate_doc`` / ``edit_doc`` take the key from ``api_key`` or the
``ANTHROPIC_API_KEY`` env var; preview / set_text / analyze are
deterministic and keyless.
"""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP


def build_local_mcp_server() -> FastMCP:
    mcp = FastMCP(
        "edit2docs",
        instructions=(
            "Generate, chat-edit, preview and text-edit Office documents "
            "(.docx, .xlsx, .pptx) on the local filesystem — the file "
            "extension picks the engine. natively editable (English-first; first-class Korean) "
            "output. Call analyze_doc first when you need addresses for "
            "set_doc_text."
        ),
    )

    from .. import simple

    @mcp.tool(
        name="generate_doc",
        description=(
            "Generate a complete document from a one-line intent; the "
            "OUTPUT extension picks the engine: .docx (Word), .xlsx "
            "(Excel, styled sheets + real numbers + formulas), .pptx "
            "(full presentation pipeline; supports template/deck_mode/"
            "pages). Optionally ground content in source documents. "
            "PPTX is slow (minutes); DOCX/XLSX take one model call."
        ),
    )
    async def generate_doc_tool(
        intent: str,
        output: str,
        sources: list[str] | None = None,
        template: str | None = None,
        deck_mode: str = "new",
        pages: tuple[int, int] = (8, 12),
        lang: str = "en-US",
        api_key: str | None = None,
    ) -> dict[str, Any]:
        result = await simple.async_generate_doc(
            intent, output=output, api_key=api_key, sources=sources,
            template=template, deck_mode=deck_mode, pages=pages, lang=lang,
        )
        return {
            "path": str(result.path),
            "page_count": result.page_count,
            "warnings": result.warnings,
        }

    @mcp.tool(
        name="edit_doc",
        description=(
            "Apply one natural-language edit turn to a local document "
            "(.docx/.xlsx/.pptx): '2번 문단 수치를 15%로', 'B3 셀을 142로', "
            "'3번 슬라이드 제목 바꿔줘'. Untouched content survives "
            "byte-identical; question-only instructions answer in `reply` "
            "without changing the file. Attach reference docs via sources."
        ),
    )
    async def edit_doc_tool(
        doc: str,
        instruction: str,
        output: str | None = None,
        sources: list[str] | None = None,
        lang: str = "en-US",
        api_key: str | None = None,
    ) -> dict[str, Any]:
        result = await simple.async_edit_doc(
            doc, instruction, output=output, api_key=api_key,
            sources=sources, lang=lang,
        )
        return {
            "path": str(result.path),
            "changed": result.changed,
            "reply": result.reply,
            "operations": result.operations,
        }

    @mcp.tool(
        name="preview_doc",
        description=(
            "Render a local document for inspection: .pptx -> slide_NNN.svg "
            "per slide; .docx/.xlsx -> preview.md (markdown). "
            "Deterministic, no LLM, no key."
        ),
    )
    async def preview_doc_tool(doc: str, out_dir: str) -> dict[str, Any]:
        rendered = simple.preview_doc(doc, out_dir=out_dir)
        if isinstance(rendered, list):
            return {"svg_paths": [str(p) for p in rendered], "page_count": len(rendered)}
        return {"preview_path": str(rendered)}

    @mcp.tool(
        name="render_doc",
        description=(
            "Render a .pptx/.docx/.xlsx to page images or a PDF via the "
            "LibreOffice-free native pipeline. to='png' -> page-1.png..N, "
            "to='pdf' -> <stem>.pdf, to='svg' -> vector pages. "
            "Deterministic, no LLM, no key."
        ),
    )
    async def render_doc_tool(
        doc: str,
        to: str = "png",
        out_dir: str | None = None,
        dpi: float = 144.0,
    ) -> dict[str, Any]:
        result = simple.render_doc(doc, to=to, out_dir=out_dir, dpi=dpi)
        return {
            "paths": [str(p) for p in result.paths],
            "page_count": result.page_count,
            "format": result.format,
            "to": result.to,
        }

    @mcp.tool(
        name="set_doc_text",
        description=(
            "Deterministic targeted edits (instant, no LLM). Untouched "
            "content is byte-preserved — charts, images, styles, merged "
            "cells and cached formulas all survive the edit. Addresses "
            "come from analyze_doc: .docx -> "
            "{action: replace|insert_after|delete, para | table/row/col, "
            "new_text|markdown}; .xlsx -> {action: set_cell|append_rows|"
            "add_sheet, sheet, cell, value, rows}; .pptx -> {slide, "
            "shape_id, para, new_text, row/col}. Prefer over edit_doc for "
            "plain value/text swaps."
        ),
    )
    async def set_doc_text_tool(
        doc: str,
        edits: list[dict],
        output: str | None = None,
    ) -> dict[str, Any]:
        result = simple.set_doc_text(doc, edits, output=output)
        return {
            "path": str(result.path),
            "applied": result.applied,
            "results": result.results,
        }

    @mcp.tool(
        name="edit_chart",
        description=(
            "Deterministically edit native charts in a local .docx/.xlsx/"
            ".pptx (instant, no LLM). Addresses come from analyze_doc's "
            "'charts' list. Edits: {chart: i, title: '...'} to retitle; or "
            "{chart: i, categories: [...], series: [{name, values: [...]}]} "
            "to set data — rewrites the chart AND its embedded workbook so "
            "Office double-click-edit matches. Byte-preserves everything else."
        ),
    )
    async def edit_chart_tool(
        doc: str,
        edits: list[dict],
        output: str | None = None,
    ) -> dict[str, Any]:
        result = simple.edit_chart(doc, edits, output=output)
        return {
            "path": str(result.path),
            "applied": result.applied,
            "results": result.results,
        }

    @mcp.tool(
        name="analyze_doc",
        description=(
            "Inspect a local document's structure and get the exact "
            "addresses set_doc_text / edit_chart need (.docx paragraph "
            "outline, .xlsx sheets + sample rows, .pptx slides + shape ids; "
            "plus a 'charts' list for edit_chart). Deterministic, no LLM, "
            "no key. Call before editing."
        ),
    )
    async def analyze_doc_tool(doc: str) -> dict[str, Any]:
        return simple.analyze_doc(doc)

    @mcp.tool(
        name="build_doc",
        description=(
            "Build a NEW document from a structured spec you wrote — "
            "DETERMINISTIC, no LLM, no key (generate_doc's engine without the "
            "model). The output extension picks the engine + `spec` shape: "
            ".docx <- markdown string; .xlsx <- {sheets: [{name, headers, "
            "rows}]}; .pptx <- {slides: [{layout, title, subtitle|bullets, "
            "notes}]} (layout in title|content|section|title_only|two_content|"
            "blank). pptx uses built-in layouts — use generate_doc for a "
            "designed deck."
        ),
    )
    async def build_doc_tool(
        spec: str | dict[str, Any],
        output: str,
        lang: str | None = None,
    ) -> dict[str, Any]:
        result = simple.build_doc(spec, output, lang=lang)
        return {
            "path": str(result.path),
            "page_count": result.page_count,
            "warnings": result.warnings,
        }

    @mcp.tool(
        name="read_doc_xml",
        description=(
            "DOCX/XLSX/PPTX are zips of XML — read that XML directly "
            "(deterministic, no LLM, no key). Without `part`: list every "
            "part in the package. With `part` (ppt/slides/slide1.xml, "
            "ppt/charts/chart1.xml, word/document.xml...): return its exact "
            "XML text. Pair with set_doc_xml to express EVERY edit OOXML "
            "can (colors, fills, fonts, geometry, chart styling...)."
        ),
    )
    async def read_doc_xml_tool(
        doc: str, part: str | None = None
    ) -> dict[str, Any]:
        if not part:
            return {"parts": simple.list_doc_parts(doc)}
        return {"part": part, "xml": simple.get_doc_xml(doc, part)}

    @mcp.tool(
        name="set_doc_xml",
        description=(
            "Patch one XML part with exact find/replace edits, or replace "
            "it whole via `xml` (deterministic, no LLM, no key). The "
            "universal escape hatch for edits the structured verbs don't "
            "cover — recolor bars/shapes, fonts, fills, geometry. `find` "
            "must match read_doc_xml's text EXACTLY (count 0 = all). The "
            "result must stay well-formed XML or nothing is written; "
            "untouched parts stay byte-identical."
        ),
    )
    async def set_doc_xml_tool(
        doc: str,
        part: str,
        edits: list[dict] | None = None,
        xml: str | None = None,
        output: str | None = None,
    ) -> dict[str, Any]:
        result = simple.set_doc_xml(doc, part, edits, xml=xml, output=output)
        return {
            "path": str(result.path),
            "applied": result.applied,
            "results": result.results,
        }

    return mcp
