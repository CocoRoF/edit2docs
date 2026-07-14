"""edit2docs as a function-calling tool set for LLM agents.

``ANTHROPIC_TOOLS`` is a ready-to-send ``tools=[...]`` list for the
Anthropic Messages API; ``run_tool`` / ``run_tool_async`` dispatch a tool
call to the library facade (:mod:`edit2docs.simple`) and return a
JSON-safe dict for the tool_result block.

Five format-dispatched verbs cover DOCX, XLSX and PPTX — the file
extension picks the engine, so agents don't juggle per-format tools.

All tools operate on local file paths. ``generate_doc`` / ``edit_doc``
need an Anthropic key (``api_key=`` on the dispatcher or
``ANTHROPIC_API_KEY``); the rest are deterministic and keyless.
"""

from __future__ import annotations

import asyncio
from typing import Any

__all__ = ["ANTHROPIC_TOOLS", "TOOL_NAMES", "run_tool", "run_tool_async"]

_DOC_PATH = {
    "type": "string",
    "description": "Path to a local document (.docx / .xlsx / .pptx).",
}

ANTHROPIC_TOOLS: list[dict[str, Any]] = [
    {
        "name": "generate_doc",
        "description": (
            "Generate a complete document from a one-line intent; the OUTPUT "
            "file extension picks the engine: .docx (Word report/proposal), "
            ".xlsx (Excel workbook with styled sheets, real numbers, "
            "formulas), .pptx (full presentation pipeline — supports "
            "template/deck_mode/pages). Optionally ground the content in "
            "source documents. English-first; Korean and any other language work equally. PPTX "
            "generation is slow (minutes); DOCX/XLSX take one model call."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "intent": {
                    "type": "string",
                    "description": "What the document is for, e.g. '3분기 실적 보고서'.",
                },
                "output": {
                    "type": "string",
                    "description": "Output path — extension (.docx/.xlsx/.pptx) selects the format.",
                },
                "sources": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional source document paths (PDF/DOCX/PPTX/XLSX/HTML).",
                },
                "template": {
                    "type": "string",
                    "description": "PPTX only: existing deck to inherit design from.",
                },
                "deck_mode": {
                    "type": "string",
                    "enum": ["new", "template_restyle", "template_extend"],
                    "description": "PPTX only: how to use the template (default new).",
                },
                "pages": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "minItems": 2,
                    "maxItems": 2,
                    "description": "PPTX only: target [min, max] page count.",
                },
                "lang": {"type": "string", "description": "BCP-47, default en-US."},
            },
            "required": ["intent", "output"],
        },
    },
    {
        "name": "build_doc",
        "description": (
            "Build a NEW document from a structured spec you already wrote — "
            "DETERMINISTIC, no LLM, no key. This is generate_doc's engine "
            "without the model: YOU produce the interchange artifact, this "
            "renders the file instantly. The OUTPUT extension picks the "
            "engine and the required `spec` shape:\n"
            "  • .docx ← `spec` is a MARKDOWN string (headings, paragraphs, "
            "lists, tables, **bold**/*italic*).\n"
            "  • .xlsx ← `spec` is an object "
            '{"sheets": [{"name","headers":[...],"rows":[[...]]}]}.\n'
            "  • .pptx ← `spec` is an object "
            '{"slides": [{"layout","title","subtitle"|"bullets","notes"}]}. '
            "layout ∈ title|content|section|title_only|two_content|blank "
            "(default content). bullets accept strings or {text,level}. "
            "pptx uses standard built-in layouts (no design pipeline — use "
            "generate_doc for that)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "spec": {
                    "type": ["string", "object"],
                    "description": (
                        "docx: markdown string. xlsx: {sheets:[...]}. "
                        "pptx: {slides:[...]}. Must match the output format."
                    ),
                },
                "output": {
                    "type": "string",
                    "description": "Output path — extension (.docx/.xlsx/.pptx) selects the format.",
                },
                "lang": {"type": "string", "description": "BCP-47, default en-US."},
            },
            "required": ["spec", "output"],
        },
    },
    {
        "name": "edit_doc",
        "description": (
            "Apply one natural-language edit turn to an existing document "
            "(.docx/.xlsx/.pptx — extension picks the engine): '2번 문단 "
            "수치를 15%로 바꿔줘', 'B3 셀을 142로', '3번 슬라이드 제목 바꿔줘'. "
            "Untouched content survives byte-identical. Question-only "
            "instructions are answered in `reply` without changing the "
            "file. Attach reference documents via `sources`."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "instruction": {"type": "string"},
                "output": {
                    "type": "string",
                    "description": "Output path (default: <input>_edited.<ext>).",
                },
                "sources": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Reference document paths for this edit.",
                },
                "lang": {"type": "string", "description": "BCP-47, default en-US."},
            },
            "required": ["doc", "instruction"],
        },
    },
    {
        "name": "render_doc",
        "description": (
            "Render/inspect a .pptx/.docx/.xlsx — LibreOffice-free native "
            "pipeline. to='png' writes page-1.png..N, to='pdf' one "
            "<stem>.pdf, to='svg' the vector pages, to='md' readable "
            "content (preview.md for docx/xlsx, per-slide SVGs for pptx). "
            "Deterministic, no LLM, no key."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "to": {
                    "type": "string",
                    "enum": ["png", "pdf", "svg", "md"],
                    "description": "Output kind (default png).",
                },
                "out_dir": {
                    "type": "string",
                    "description": "Output directory (default <doc dir>/render).",
                },
                "dpi": {
                    "type": "number",
                    "description": "Raster resolution (default 144).",
                },
            },
            "required": ["doc"],
        },
    },
    {
        "name": "set_doc_text",
        "description": (
            "Deterministic structured edits — text AND charts in one call "
            "(instant, no LLM). Untouched content is byte-preserved. "
            "Addresses come from analyze_doc: .docx -> "
            "{action: replace|insert_after|delete, para | table/row/col, "
            "new_text|markdown}; .xlsx -> {action: set_cell|append_rows|"
            "add_sheet, sheet, cell, value, rows}; .pptx -> {slide, "
            "shape_id, para, new_text, row/col for tables}. CHART edits "
            "(any format) carry a `chart` index from analyze_doc's "
            "'charts' list: {chart: i, title: '...'} retitles; {chart: i, "
            "categories: [...], series: [{name, values: [...]}]} sets the "
            "data AND the embedded workbook (Office double-click-edit "
            "matches). Prefer this over edit_doc for value/text/data swaps."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "edits": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": (
                        "Edit objects; ones with a `chart` key address "
                        "charts, the rest are format-specific text edits."
                    ),
                },
                "output": {
                    "type": "string",
                    "description": "Output path (default: <input>_edited.<ext>).",
                },
            },
            "required": ["doc", "edits"],
        },
    },
    {
        "name": "read_doc_xml",
        "description": (
            "DOCX/XLSX/PPTX are zips of XML — this reads that XML directly. "
            "Deterministic, no LLM, no key. Without `part`: list every part "
            "in the package (slides, charts, styles, themes, sheets...). "
            "With `part` (e.g. ppt/slides/slide1.xml, ppt/charts/chart1.xml, "
            "word/document.xml): return that part's exact XML text. Read the "
            "XML, copy exact substrings, then patch them with set_doc_xml — "
            "together they express EVERY edit OOXML can (colors, fills, "
            "fonts, geometry, chart styling...) without python-pptx."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "part": {
                    "type": "string",
                    "description": "Part name to read. Omit to list all parts.",
                },
            },
            "required": ["doc"],
        },
    },
    {
        "name": "set_doc_xml",
        "description": (
            "Patch, CREATE or DELETE one XML part. Deterministic, no LLM, "
            "no key — the universal escape hatch for everything the "
            "structured verbs don't cover: recolor bars/shapes, fonts, "
            "fills, geometry, add/remove slides. `edits` patches an "
            "existing part (`find` must match read_doc_xml's text EXACTLY, "
            "count 0 = all). `xml` replaces the whole part — and CREATES "
            "it if missing (pass `content_type` to register the new part, "
            "e.g. adding slideN.xml + its _rels/*.rels, then patching "
            "presentation.xml + its rels = a new slide). `delete: true` "
            "removes the part. The result must stay well-formed XML or "
            "nothing is written; untouched parts stay byte-identical."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "part": {
                    "type": "string",
                    "description": "Part to patch/create/delete, e.g. ppt/charts/chart1.xml.",
                },
                "edits": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "find": {"type": "string"},
                            "replace": {"type": "string"},
                            "count": {"type": "integer"},
                        },
                        "required": ["find", "replace"],
                    },
                    "description": "Exact-substring edits (one mode: edits | xml | delete).",
                },
                "xml": {
                    "type": "string",
                    "description": "Full part XML — replaces, or creates a missing part.",
                },
                "content_type": {
                    "type": "string",
                    "description": (
                        "[Content_Types].xml Override for a NEWLY created part, "
                        "e.g. application/vnd.openxmlformats-officedocument."
                        "presentationml.slide+xml"
                    ),
                },
                "delete": {
                    "type": "boolean",
                    "description": "Remove the part (also patch referencing rels).",
                },
                "output": {
                    "type": "string",
                    "description": "Output path (default: <input>_edited.<ext>).",
                },
            },
            "required": ["doc", "part"],
        },
    },
    {
        "name": "analyze_doc",
        "description": (
            "Inspect a document's structure and get the exact addresses "
            "set_doc_text needs: .docx -> paragraph/table-cell outline "
            "with indices; .xlsx -> sheets, dimensions, sample rows; "
            ".pptx -> slides, theme, per-paragraph shape ids. Every format "
            "also returns a 'charts' list (kinds/titles/series) for chart "
            "edits. Deterministic, no LLM, no key. Call before editing."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"doc": _DOC_PATH},
            "required": ["doc"],
        },
    },
]

TOOL_NAMES = [t["name"] for t in ANTHROPIC_TOOLS]


async def run_tool_async(
    name: str, tool_input: dict[str, Any], *, api_key: str | None = None
) -> dict[str, Any]:
    """Dispatch one tool call; returns a JSON-safe result dict."""
    from . import simple

    args = dict(tool_input)
    if name == "generate_doc":
        pages_raw = args.pop("pages", None)
        pages = tuple(pages_raw) if pages_raw and len(pages_raw) == 2 else (8, 12)
        result = await simple.async_generate_doc(
            args.pop("intent"),
            output=args.pop("output"),
            api_key=api_key,
            sources=args.pop("sources", None),
            template=args.pop("template", None),
            deck_mode=args.pop("deck_mode", "new") or "new",
            pages=pages,  # type: ignore[arg-type]
            lang=args.pop("lang", "en-US") or "en-US",
        )
        return {
            "path": str(result.path),
            "page_count": result.page_count,
            "warnings": result.warnings,
        }
    if name == "edit_doc":
        result = await simple.async_edit_doc(
            args.pop("doc"),
            args.pop("instruction"),
            output=args.pop("output", None),
            api_key=api_key,
            sources=args.pop("sources", None),
            lang=args.pop("lang", "en-US"),
        )
        return {
            "path": str(result.path),
            "changed": result.changed,
            "reply": result.reply,
            "operations": result.operations,
        }
    if name == "render_doc":
        result = simple.render_doc(
            args["doc"],
            to=args.get("to", "png"),
            out_dir=args.get("out_dir"),
            dpi=float(args.get("dpi", 144.0)),
        )
        return {
            "paths": [str(p) for p in result.paths],
            "page_count": result.page_count,
            "format": result.format,
            "to": result.to,
        }

    if name == "set_doc_text":
        # One structured-edit surface: dicts with a `chart` key go to the
        # chart engine, the rest to the text engine, chained on one output.
        all_edits = list(args["edits"])
        text_edits = [e for e in all_edits if "chart" not in e]
        chart_edits = [e for e in all_edits if "chart" in e]
        output = args.get("output")
        path, applied, results = args["doc"], 0, []
        if text_edits:
            r = simple.set_doc_text(path, text_edits, output=output)
            path, applied = str(r.path), applied + r.applied
            results.extend(r.results)
        if chart_edits:
            # Continue on the text-edit output when chaining; otherwise let
            # edit_chart use its own default (<input>_chart.<ext>).
            chart_out = output if output is not None else (
                path if path != str(args["doc"]) else None
            )
            r = simple.edit_chart(path, chart_edits, output=chart_out)
            path, applied = str(r.path), applied + r.applied
            results.extend(r.results)
        return {"path": path, "applied": applied, "results": results}
    if name == "analyze_doc":
        return simple.analyze_doc(args["doc"])
    if name == "read_doc_xml":
        part = args.get("part")
        if not part:
            return {"parts": simple.list_doc_parts(args["doc"])}
        return {"part": part, "xml": simple.get_doc_xml(args["doc"], part)}
    if name == "set_doc_xml":
        result = simple.set_doc_xml(
            args["doc"],
            args["part"],
            args.get("edits"),
            xml=args.get("xml"),
            content_type=args.get("content_type"),
            delete=bool(args.get("delete")),
            output=args.get("output"),
        )
        return {
            "path": str(result.path),
            "applied": result.applied,
            "results": result.results,
        }
    if name == "build_doc":
        result = simple.build_doc(
            args["spec"], args["output"], lang=args.get("lang")
        )
        return {
            "path": str(result.path),
            "page_count": result.page_count,
            "warnings": result.warnings,
        }
    raise ValueError(f"unknown edit2docs tool: {name!r} (known: {TOOL_NAMES})")


def run_tool(
    name: str, tool_input: dict[str, Any], *, api_key: str | None = None
) -> dict[str, Any]:
    """Sync wrapper for :func:`run_tool_async`."""
    return asyncio.run(run_tool_async(name, tool_input, api_key=api_key))
