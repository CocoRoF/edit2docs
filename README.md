# edit2docs

**AI-agent-native document engine: generate and chat-edit DOCX, XLSX and PPTX as a Python library, agent tool set, MCP server or hosted service. English-first, with first-class Korean support.**

[![PyPI](https://img.shields.io/pypi/v/edit2docs)](https://pypi.org/project/edit2docs/)
[![Python](https://img.shields.io/badge/python-3.12%2B-blue)](https://pypi.org/project/edit2docs/)
[![License: Apache-2.0](https://img.shields.io/badge/license-Apache--2.0-green)](./LICENSE)

[한국어 README](./README.ko.md)

## What it is

LLM agents that "edit a Word/Excel/PowerPoint file" usually either regenerate
the whole file or flatten it into pictures, destroying charts, tables,
styles and formulas on the way. `edit2docs` gives an agent (or a script) a
small set of **format-dispatched verbs** that work on real OOXML:

* **Generate** a complete document from a one-line intent (LLM) or from a spec
  you write (deterministic, no key).
* **Edit** an existing file by chat (LLM) or by exact addresses
  (deterministic) — untouched package parts stay **byte-identical**, so a
  chart, sparkline, image or cached formula is never collateral damage.
* **Inspect and render** — an addressable outline, page PNG/PDF/SVG, markdown.

The output is always natively editable (real paragraphs, cells, charts).
One engine, four surfaces: Python library, agent tools (function calling),
a local stdio MCP server, and a hosted FastAPI service.

## Install

Requires Python 3.12+.

```bash
pip install edit2docs              # library + agent tools + local MCP server
pip install "edit2docs[server]"    # + the hosted multi-tenant service
```

Optional extras: `images` (Gemini/OpenAI image backends), `svg-fallback`
(cairosvg), `dev` (tests/lint, includes `server`). The repo's `main` branch is
ahead of PyPI at times (see [Version history](#version-history)); to get it:
`pip install "edit2docs @ git+https://github.com/CocoRoF/edit2docs"`.

## Quick start

Deterministic verbs need no API key. This runs as-is:

```python
from edit2docs import build_doc, analyze_doc, set_doc_text, render_doc

build_doc({"slides": [
    {"layout": "title", "title": "Q3 Review"},
    {"layout": "content", "title": "Highlights", "bullets": ["Revenue +12%", "Churn down"]},
]}, "deck.pptx")

info = analyze_doc("deck.pptx")      # outline with exact addresses
title = next(t for t in info["slides"][1]["texts"] if t["text"] == "Highlights")
# title == {"para": 0, "text": "Highlights", "shape_id": 2}

r = set_doc_text("deck.pptx", [
    {"slide": 1, "shape_id": title["shape_id"], "para": title["para"], "new_text": "Key results"},
])
print(r.path, r.applied)             # deck_edited.pptx 1  (the input is not overwritten)

render_doc("deck_edited.pptx", to="png", out_dir="pages")   # page PNGs, no LibreOffice
```

Generative verbs need an Anthropic key (`api_key=` or `ANTHROPIC_API_KEY`):

```python
from edit2docs import generate_doc, edit_doc

generate_doc("Executive briefing on Q3 sales", output="deck.pptx")
r = edit_doc("deck.pptx", "Make slide 3's title more assertive")
print(r.reply)        # the editor explains what it changed
print(r.path)         # deck_edited.pptx (pass output=... to choose)
```

## The verbs

The file extension picks the engine. All of these exist as library functions
in `edit2docs`; `[det]` verbs are deterministic and keyless, `[LLM]` verbs are
BYOK.

| verb | what it does | |
|---|---|---|
| `build_doc(spec, output)` | render a document from a spec you write: markdown (docx), `{"sheets": [...]}` (xlsx), `{"slides": [...], "theme": {...}}` (pptx, incl. themed decks) | det |
| `generate_doc(intent, *, output, ...)` | intent (+ optional `sources`, PPTX `template`) → designed document | LLM |
| `edit_doc(doc, instruction, ...)` | one natural-language edit turn; untouched content stays byte-identical | LLM |
| `analyze_doc(doc)` | structure outline with the exact addresses the edit verbs take, plus a `charts` list | det |
| `set_doc_text(doc, edits, *, output=None)` | targeted edits: docx paragraphs/table cells, xlsx cells/rows/sheets, pptx shape paragraphs/table cells, and native chart title/data (edits with a `chart` index) | det |
| `edit_chart(doc, edits)` / `list_charts(doc)` | edit a native chart's data or title (chart XML and embedded workbook stay in sync), docx/xlsx/pptx | det |
| `arrange_doc(doc, ops, *, output=None)` | structure: duplicate / move / delete whole slides (pptx) and sheets (xlsx), rename sheets | det |
| `list_doc_parts(doc)` / `get_doc_xml(doc, part)` / `set_doc_xml(doc, part, ...)` | read the package part map or one part's XML; patch, create or delete a part (colors, fonts, geometry, anything) | det |
| `preview_doc(doc, *, out_dir=None)` | pptx → per-slide SVG, docx/xlsx → markdown | det |
| `render_doc(doc, *, to="png", out_dir=None, dpi=144.0)` | `to` = `png` / `pdf` / `svg` / `md`; resvg + PyMuPDF, no LibreOffice | det |
| `doc_guide(topic=None)` | progressive-disclosure guide for agents (topics: build, generate, edit, edit.text, edit.chart, arrange, edit.xml, recipes.slides, recipes.colors, render) | det |

Async variants: `async_generate_doc`, `async_edit_doc`. PPTX-specific
aliases (`generate_pptx`, `edit_pptx`, `preview_pptx`, `set_pptx_text`,
`analyze_pptx`) also exist.

Result objects: `GenerateResult(path, page_count, design_spec, warnings)`,
`EditResult` (has `.reply`, `.operations`, `.path`), `TextEditsResult(path,
applied, results)` (per-edit `status`: `applied | stale | not_found |
invalid`), `ArrangeResult`, `RenderResult(paths, page_count, format, to)`.
Verbs that write a derived file default to `<stem>_edited.<ext>`
(`_chart`, `_arranged` for those verbs); pass `output=` to choose.

### Generate

```python
generate_doc("Q3 performance report", output="report.docx")
generate_doc("Quarterly sales summary", output="sales.xlsx", sources=["raw.pdf"])
generate_doc("Executive briefing on Q3 sales", output="deck.pptx",
             template="brand.pptx",          # optional user PPTX template
             deck_mode="template_restyle",   # "new" | "template_restyle" | "template_extend"
             pages=(8, 12))                  # target page range (pptx)
generate_doc("3분기 실적 보고서", output="report.docx", lang="ko-KR")
```

`sources` accepts PDF / DOCX / DOC / PPTX / XLSX / HTML / EPUB / IPYNB paths;
each is converted to markdown and given to the writer as reference material.
Full signature: `generate_doc(intent, *, output, api_key=None, sources=None,
template=None, deck_mode="new", pages=(8, 12), lang="en-US", model=...)`.

### Edit by chat

```python
r = edit_doc("report.docx", "Add a 'deployment complete' item to the progress section")
print(r.reply, r.operations)

r = edit_doc("deck.pptx", "이 문서 내용을 반영해서 3번 슬라이드를 고쳐줘",
             sources=["notes.pdf"], lang="ko-KR",
             chat_history=[{"role": "user", "content": "..."},
                           {"role": "assistant", "content": "..."}])
```

The planner sees a numbered outline of the document, plans the minimal
operations, and a deterministic engine applies them. If planning fails the
reply says so instead of silently doing nothing. For decks, targeted edits
run as in-place, byte-preserving operations on the real shapes (25 ops, e.g.
`set_text`, `set_runs`, `set_table_cell`, `insert_row`, `set_chart_data`,
`add_slide`, `set_notes`, `set_hyperlink`, `set_theme_color`); new-slide and
full-redesign turns that need it fall back to regenerating the slide as SVG.
The op catalog is introspectable (see below).

### Edit deterministically

```python
set_doc_text("report.docx", [
    {"action": "replace", "para": 0, "new_text": "Q3 Final Report"},
    {"action": "insert_after", "para": 0, "markdown": "New paragraph"},
    {"action": "replace", "table": 0, "row": 1, "col": 2, "new_text": "142"},
    {"action": "delete", "para": 5},
])
set_doc_text("sales.xlsx", [
    {"action": "set_cell", "sheet": "Sales", "cell": "B3", "value": 142},
    {"action": "append_rows", "sheet": "Sales", "rows": [["Q4", 160]]},
    {"action": "add_sheet", "sheet": "Notes", "headers": ["a"], "rows": [["x"]]},
])
set_doc_text("deck.pptx", [{"slide": 1, "shape_id": 5, "para": 0, "new_text": "New title"}])

edit_chart("deck.pptx", [
    {"chart": 0, "title": "Q3 Sales"},
    {"chart": 0, "categories": ["Q1", "Q2", "Q3"],
     "series": [{"name": "Sales", "values": [120, 135, 150]}]},
])

arrange_doc("deck.pptx", [{"op": "duplicate", "target": 0}, {"op": "move", "target": 2, "to": 0}])
```

Optional `old_text` / `old_value` guards reject stale edits. Take addresses
from `analyze_doc`. Run `doc_guide("edit.text")` for the full shapes.

The PPTX surgical engine is also directly callable:

```python
from edit2docs.documents.pptx_engine import apply_pptx_edits, describe_ops, find_shapes

describe_ops()                                    # machine-readable op catalog (address + payload per op)
find_shapes(pptx_bytes, text="Revenue")           # query shapes by text/kind/name/slide
new_bytes, results = apply_pptx_edits(pptx_bytes, edits, dry_run=True)   # validate without writing
# atomic=True -> all-or-nothing
```

## Agent tools (function calling)

The verbs `doc_guide`, `analyze_doc`, `render_doc`, `set_doc_text`,
`arrange_doc`, `read_doc_xml`, `set_doc_xml`, `build_doc`, `generate_doc` and
`edit_doc` are exposed as tool schemas plus a dispatcher:

```python
import anthropic
from edit2docs.agent_tools import ANTHROPIC_TOOLS, run_tool

client = anthropic.Anthropic()
msg = client.messages.create(
    model="claude-sonnet-5",
    max_tokens=2048,
    tools=ANTHROPIC_TOOLS,
    messages=[{"role": "user", "content": "Fix the title of slide 3 in deck.pptx"}],
)
for block in msg.content:
    if block.type == "tool_use":
        result = run_tool(block.name, block.input)   # run_tool_async also exists
```

`OPENAI_TOOLS` is the function-calling shape for OpenAI-style APIs.
`tool_specs(provider, extension="xlsx")` returns only the verbs (with
format-specialized descriptions) relevant to that file type; the unscoped
default returns the full set. Chart editing goes through `set_doc_text`
(edits with a `chart` index) on this surface.

## Local MCP server (zero infra)

`pip install edit2docs` installs an `edit2docs-mcp` stdio server exposing the
same ten verbs over local files.

```jsonc
// Claude Desktop / Claude Code / Cursor
{
  "mcpServers": {
    "edit2docs": {
      "command": "edit2docs-mcp",
      "env": { "ANTHROPIC_API_KEY": "sk-ant-..." }   // only generative tools need it
    }
  }
}
```

Without installing: `uvx --from edit2docs edit2docs-mcp`. See
[docs/mcp-clients.md](./docs/mcp-clients.md) for client setups (note: that
page describes the hosted server's tool list, which differs from the local
one).

## Hosted service

```bash
pip install "edit2docs[server]"
edit2docs serve [--host H] [--port P] [--reload]    # FastAPI, default port 8000
```

Standalone mode needs no external infrastructure: SQLite + local-filesystem
storage + an inline job queue under `EDIT2DOCS_DATA_DIR`. Postgres, Redis
(arq worker queue) and S3-compatible storage are enabled by env vars.

| endpoint | purpose |
|---|---|
| `POST /v1/assets`, `GET/DELETE /v1/assets/{id}` | upload / fetch / delete documents (200 MB default cap) |
| `POST /v1/jobs/generate-deck` | queue a generation job; `output_format` = `pptx` / `docx` / `xlsx` |
| `POST /v1/jobs/edit-deck` | queue a chat-edit job |
| `GET /v1/jobs/{id}`, `GET /v1/jobs/{id}/events` | status; SSE progress stream (stages plus per-operation live-edit events with addressable targets) |
| `POST /v1/preview` | pptx → per-slide SVGs, docx/xlsx → addressable HTML (`data-e2d-*`) |
| `POST /v1/text-edits` | deterministic targeted edits |
| `GET /v1/models` | live model list (proxies the Anthropic Models API with the caller's key; curated fallback) |
| `GET /health` | liveness and mode report |
| `/mcp`, `/mcp-sse` | MCP over Streamable HTTP / SSE (deck-job oriented tools: `generate_deck`, `edit_deck`, `upload_source`, `request_upload_url`, `get_asset`, `download_url`, `list_templates`, `list_voices`, `hello`) |

Anthropic keys are BYOK per request (`X-Anthropic-API-Key` header). Errors are
bilingual: `message` follows `Accept-Language`, with English and Korean
variants available.

### Configuration

Environment variables, prefix `EDIT2DOCS_` (see `src/edit2docs/config.py`):

| var | default | notes |
|---|---|---|
| `EDIT2DOCS_DEFAULT_LANG` | `en-US` | `ko-KR` makes a deployment Korean-by-default |
| `EDIT2DOCS_HOST` / `EDIT2DOCS_PORT` | `0.0.0.0` / `8000` | overridable by `serve` flags |
| `EDIT2DOCS_DATA_DIR` | `/data/edit2docs` | SQLite + file storage root |
| `EDIT2DOCS_DATABASE_URL` | SQLite under data dir | e.g. `postgresql+asyncpg://...` |
| `EDIT2DOCS_REDIS_URL` | unset (inline queue) | enables the arq worker queue |
| `EDIT2DOCS_S3_ENDPOINT_URL` + `EDIT2DOCS_S3_BUCKET` | unset (local fs) | with `S3_REGION`, `S3_ACCESS_KEY_ID`, `S3_SECRET_ACCESS_KEY`, `S3_PUBLIC_BASE_URL` |
| `EDIT2DOCS_RAW_SIGNING_KEY` | random per start | set for production so signed local URLs survive restarts |
| `EDIT2DOCS_AUTH_DEV_API_KEY` | unset (anonymous) | single bearer token for small deployments |
| `EDIT2DOCS_MAX_UPLOAD_SIZE_BYTES` | 200 MB | match your reverse proxy |
| `EDIT2DOCS_MODEL_{PLANNER,WRITER,STRATEGIST,EXECUTOR}` | request model | per-role model override |
| `EDIT2DOCS_STRATEGIST_SOURCE_CHAR_CAP` | `60000` | per-source cap fed to the deck strategist (0 = uncapped) |

Image generation backends (`images` extra) read provider variables such as
`IMAGE_BACKEND`, `OPENAI_API_KEY`, `GEMINI_API_KEY`; see `.env.example`.

### Web studio

[**edit2docs-web**](https://github.com/CocoRoF/edit2docs-web) is the official
frontend: upload, generation with staged SSE progress, and a chat-edit studio
whose canvas highlights the exact paragraph / cell / slide each operation
touches. Point it at the engine with `EDIT2DOCS_SERVER_INTERNAL_URL` and
`EDIT2DOCS_SERVER_API_KEY`.

## How each format works

* **DOCX** — the writer LLM emits constrained markdown; a deterministic
  python-docx renderer produces the file. Edits are paragraph/table-cell
  addressed. Preview is native addressable HTML (`data-e2d-para`,
  `data-e2d-table`, `data-e2d-cell`).
* **XLSX** — the designer LLM emits a YAML sheet spec; openpyxl renders it.
  Edits are `set_cell` / `append_rows` / `add_sheet` with staleness guards.
  Preview is a grid with `data-e2d-cell="B3"` addresses and cached formula
  values.
* **PPTX** — strategist → per-page SVG → native DrawingML, user-template
  restyle/extend, optional Edge-TTS narration. Chat edits use the surgical
  in-place engine described above.

Lossless editing is built on [contextifier](https://github.com/CocoRoF/Contextifier)'s
raw OOXML layer (`contextifier>=0.8.0`): only the parts an edit touches are
rewritten.

### Native charts and tables (PPTX)

SVG groups marked `data-pptx-native="chart|table"` export as real PowerPoint
charts (chart part plus embedded workbook) or native `<a:tbl>` tables:

```xml
<g id="sales_chart" data-pptx-native="chart">
  <metadata data-pptx-native="chart">
    { "name": "sales_chart", "x": 125, "y": 141, "width": 1000, "height": 440,
      "type": "bar", "categories": ["Q1", "Q2", "Q3"],
      "series": [{ "name": "Sales", "values": [120, 135, 150] }] }
  </metadata>
  <!-- fallback shapes, used when native export is off -->
</g>
```

Opt in with `ExportRequest(native_objects=True)` (`tools/export.py`).

## Languages

English is the default (`lang="en-US"`). Korean is first-class: Hangul-aware
text widths, per-run OOXML `lang` from the actual script, Korean font stacks,
a complete message catalog and localized replies. Switch per call
(`lang="ko-KR"`), per request (`Accept-Language`), or per deployment
(`EDIT2DOCS_DEFAULT_LANG=ko-KR`). zh-CN / zh-TW / ja-JP get the same script
detection and font-stack handling.

## Ecosystem

| Repo | What it is |
|---|---|
| [edit2docs-web](https://github.com/CocoRoF/edit2docs-web) | Web studio for the hosted service (Next.js) |
| [contextifier](https://github.com/CocoRoF/Contextifier) | Raw OOXML layer used for lossless edits |
| [ppt-master](https://github.com/hugohe3/ppt-master) | Upstream (MIT) the PPTX core under `src/edit2docs/core/` is derived from |
| [edit2ppt](https://github.com/CocoRoF/edit2ppt) | Sister project; the deck pipeline originates there |

## Project layout

`src/edit2docs/`: `simple.py` (library verbs) · `agent_tools.py`,
`agent_guide.py`, `tool_matrix.py` (agent surface) · `documents/` (docx/xlsx/pptx
engines, arrange, chart and XML editing) · `tools/` (LLM pipelines: generate,
edit, preview, export) · `core/` (ppt-master-derived deck engine) · `mcp/`
(local stdio + hosted servers) · `api/` (FastAPI) · `render/`, `i18n/`,
`llm/`, `db/`, `storage/`, `workers/`, `services/`. Design notes live in
[docs/](./docs/) and [PPTX_EDITING_UPGRADE_PLAN.md](./PPTX_EDITING_UPGRADE_PLAN.md).

## Development

```bash
git clone https://github.com/CocoRoF/edit2docs && cd edit2docs
uv venv .venv && uv pip install -e ".[server,dev]"
.venv/bin/python -m pytest tests/          # 949 passed, 1 skipped at 0.19.0
.venv/bin/python -m ruff check src/edit2docs
```

`scripts/lint_ascii_paths.py` enforces ASCII-only paths (run by the tests).

## Version history

Current version in the repo: **0.19.0**. As of 2026-10-01 PyPI has 0.16.1;
0.17.0-0.19.0 are on `main` only.

| version | highlights |
|---|---|
| 0.19.0 | native, layout-aware `add_slide`; ops `set_notes`, `set_bullet`, `set_hyperlink`, `set_z_order`, `set_legend`, `set_series_color`, `set_theme_color`, `set_theme_font` (25 PPTX ops); needs `contextifier>=0.8.0` |
| 0.18.0 | self-describing PPTX engine: `OP_CATALOG` / `describe_ops()`, `dry_run`, `atomic`, `find_shapes`; ops `set_runs`, `add_textbox`, `delete_shape`, `duplicate_shape`, `insert_column`, `delete_column`, `merge_cells` |
| 0.17.x | structural PPTX editing: addressable, byte-preserving in-place edits instead of whole-slide SVG regeneration for targeted turns |
| 0.16.x | extension-scoped tools (`tool_specs(extension=...)`, `tool_matrix`); `GET /v1/models`; hardening of slide-edit retries and failed-job handling |
| 0.15.1 | `arrange_doc` (0.15.0 shipped without its implementation; use 0.15.1+) |
| 0.14.1 | cap `mcp<2.0` (2.0 removed `mcp.server.fastmcp`) |
| 0.10-0.14 | `build_doc`, `read_doc_xml` / `set_doc_xml`, hierarchical `doc_guide`, themed decks |
| 0.9.0 | token optimization: prompt-cache restructuring, input caps, per-role model tiering |
| 0.8.0 | lossless editing on contextifier; `edit_chart` |
| 0.7.x | upstream sync with ppt-master v3.1, native chart/table export, English-first, Apache-2.0 relicense (0.7.1) |
| 0.1-0.6 | multi-format engine, hosted API, live edit streaming, addressable previews, `render_doc` |

## License

[Apache License 2.0](./LICENSE). The PPTX core under `src/edit2docs/core/` is
derived from [ppt-master](https://github.com/hugohe3/ppt-master) (MIT,
Copyright (c) Hugo He); its original MIT terms are preserved in
[NOTICE](./NOTICE) and [LICENSE.ppt-master.MIT](./LICENSE.ppt-master.MIT).
