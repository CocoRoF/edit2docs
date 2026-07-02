# edit2docs

**AI-agent-native document engine — DOCX · XLSX · PPTX. Korean-first.**

[한국어 README](./README.ko.md) · Sister project of
[edit2ppt](https://github.com/CocoRoF/edit2ppt) (the PPTX pipeline is inherited
from it) · Built on [ppt-master](https://github.com/hugohe3/ppt-master) (MIT)

---

`edit2docs` generates complete Office documents from a one-line intent and
chat-edits existing files — Word reports, Excel workbooks and PowerPoint decks —
always producing natively editable OOXML. One engine, four surfaces:

```bash
pip install edit2docs              # library + agent tools + local MCP
pip install "edit2docs[server]"    # + the hosted multi-tenant service
```

## The five verbs

Every surface exposes the same five format-dispatched verbs — the file
extension picks the engine:

| verb | what it does | LLM? |
|---|---|---|
| `generate_doc` | intent (+ optional sources / PPTX template) → complete document | ✳ |
| `edit_doc` | one natural-language edit turn; untouched content stays byte-identical | ✳ |
| `preview_doc` | .pptx → per-slide SVG · .docx/.xlsx → markdown | — |
| `analyze_doc` | structure outline **with the exact addresses `set_doc_text` needs** | — |
| `set_doc_text` | deterministic targeted edits (paragraphs / cells / slide text) | — |

## 1 · Python library

```python
from edit2docs import generate_doc, edit_doc, preview_doc, analyze_doc, set_doc_text

generate_doc("3분기 실적 보고서", output="report.docx")
generate_doc("분기별 매출 정리", output="sales.xlsx", sources=["raw.pdf"])
generate_doc("Q3 영업 결과 임원 보고", output="deck.pptx", template="brand.pptx",
             deck_mode="template_restyle")

r = edit_doc("report.docx", "진행 사항 섹션에 배포 완료 항목을 추가해줘")
print(r.reply, r.operations)

info = analyze_doc("sales.xlsx")               # sheets + sample rows
set_doc_text("sales.xlsx", [{"sheet": "매출", "cell": "B3", "value": 142}])
```

BYOK: `api_key=...` or `ANTHROPIC_API_KEY`. The deterministic verbs need no key.
Async variants: `async_generate_doc`, `async_edit_doc`.

## 2 · Agent tools (function calling)

```python
from edit2docs.agent_tools import ANTHROPIC_TOOLS, run_tool

msg = client.messages.create(model="claude-opus-4-7", tools=ANTHROPIC_TOOLS, ...)
for block in msg.content:
    if block.type == "tool_use":
        result = run_tool(block.name, block.input)
```

## 3 · Local MCP server (zero infra)

```jsonc
// Claude Desktop / Claude Code / Cursor
{ "mcpServers": { "edit2docs": {
    "command": "edit2docs-mcp",
    "env": { "ANTHROPIC_API_KEY": "sk-ant-..." }
} } }
```

## 4 · Hosted service

`pip install "edit2docs[server]"` then `edit2docs serve` — the FastAPI service
(REST + SSE jobs + hosted MCP) inherited from edit2ppt: multi-tenant assets,
job queue, S3-compatible storage, PPTX studio endpoints.

## How each format works

* **DOCX** — the writer LLM emits a constrained markdown document; a
  deterministic renderer (python-docx) turns it into styled Word. Edits are
  paragraph-addressed operations (`replace` / `insert_after` / `delete`,
  table cells by row/col) so untouched paragraphs keep their formatting.
* **XLSX** — the designer LLM emits a YAML *sheet spec* (sheets / headers /
  rows / number formats, formulas allowed); openpyxl renders styled sheets.
  Edits are `set_cell` / `append_rows` / `add_sheet` with staleness guards.
* **PPTX** — the full multi-stage deck pipeline inherited from edit2ppt:
  strategist → per-page SVG → native DrawingML, user-PPTX templates
  (restyle/extend), chat-edit with slide recompose, per-paragraph text edits
  incl. table cells.

Every LLM planner follows the same contract: fenced `reply` + `edit_plan`
blocks, one retry with a format reminder, and an honest reply (instead of a
silent no-op) when planning fails. Korean-native throughout: Hangul-aware
widths, `lang="ko-KR"` OOXML runs, bilingual errors.

## License

MIT. PPTX core forked from ppt-master (MIT) via edit2ppt.
