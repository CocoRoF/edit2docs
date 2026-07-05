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
            "source documents. Korean-first; any language works. PPTX "
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
                "lang": {"type": "string", "description": "BCP-47, default ko-KR."},
            },
            "required": ["intent", "output"],
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
                "lang": {"type": "string", "description": "BCP-47, default ko-KR."},
            },
            "required": ["doc", "instruction"],
        },
    },
    {
        "name": "preview_doc",
        "description": (
            "Render a document for inspection: .pptx -> one self-contained "
            "SVG file per slide; .docx/.xlsx -> a markdown rendering "
            "(preview.md). Deterministic, no LLM, no key."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "out_dir": {
                    "type": "string",
                    "description": "Directory for the preview file(s).",
                },
            },
            "required": ["doc", "out_dir"],
        },
    },
    {
        "name": "render_doc",
        "description": (
            "Render a .pptx/.docx/.xlsx to page images or a PDF — the "
            "LibreOffice-free native pipeline (per-page SVG -> resvg PNG "
            "-> PDF). to='png' writes page-1.png..N, to='pdf' one "
            "<stem>.pdf, to='svg' the vector pages. Deterministic, no "
            "LLM, no key."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "to": {
                    "type": "string",
                    "enum": ["png", "pdf", "svg"],
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
            "Deterministic targeted edits (instant, no LLM, formatting "
            "preserved). Addresses come from analyze_doc: .docx -> "
            "{action: replace|insert_after|delete, para | table/row/col, "
            "new_text|markdown}; .xlsx -> {action: set_cell|append_rows|"
            "add_sheet, sheet, cell, value, rows}; .pptx -> {slide, "
            "shape_id, para, new_text, row/col for tables}. Prefer this "
            "over edit_doc for plain value/text swaps."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "doc": _DOC_PATH,
                "edits": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": "Format-specific edit objects (see tool description).",
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
        "name": "analyze_doc",
        "description": (
            "Inspect a document's structure and get the exact addresses "
            "set_doc_text needs: .docx -> paragraph/table-cell outline with "
            "indices; .xlsx -> sheets, dimensions, sample rows; .pptx -> "
            "slides, theme, per-paragraph shape ids. Deterministic, no LLM, "
            "no key. Call this before editing."
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
            lang=args.pop("lang", "ko-KR") or "ko-KR",
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
            lang=args.pop("lang", "ko-KR"),
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

    if name == "preview_doc":
        rendered = simple.preview_doc(args["doc"], out_dir=args["out_dir"])
        if isinstance(rendered, list):
            return {"svg_paths": [str(p) for p in rendered], "page_count": len(rendered)}
        return {"preview_path": str(rendered)}
    if name == "set_doc_text":
        result = simple.set_doc_text(
            args["doc"], args["edits"], output=args.get("output")
        )
        return {
            "path": str(result.path),
            "applied": result.applied,
            "results": result.results,
        }
    if name == "analyze_doc":
        return simple.analyze_doc(args["doc"])
    raise ValueError(f"unknown edit2docs tool: {name!r} (known: {TOOL_NAMES})")


def run_tool(
    name: str, tool_input: dict[str, Any], *, api_key: str | None = None
) -> dict[str, Any]:
    """Sync wrapper for :func:`run_tool_async`."""
    return asyncio.run(run_tool_async(name, tool_input, api_key=api_key))
