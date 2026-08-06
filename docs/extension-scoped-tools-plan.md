# Extension-scoped tool exposure — design report

**Status:** proposal (for review before implementation)
**Date:** 2026-08-06
**Goal:** surface only the tools (and only the tool *detail*) appropriate to the
extension of the file the user provides — so an agent editing a `.xlsx` never sees
slide ops, and one editing a `.docx` never sees `arrange_doc` at all.

---

## Part 1 — Complete tool inventory (what edit2docs provides today)

### The agent verb surface (9 verbs — `ANTHROPIC_TOOLS` / `OPENAI_TOOLS`, mirrored to local MCP and geny's `Doc*`)

| # | verb | family | key? | .docx | .xlsx | .pptx | per-format difference |
|---|---|---|---|:--:|:--:|:--:|---|
| 1 | `doc_guide` | INSPECT | no | ✅ | ✅ | ✅ | the skill map + topic guides; **content** is format-specific |
| 2 | `analyze_doc` | INSPECT | no | ✅ | ✅ | ✅ | output key differs: docx→`outline[]`, xlsx→`sheets[]`, pptx→`slides[]` (+`charts[]` always) |
| 3 | `render_doc` | INSPECT | no | ✅ | ✅ | ✅ | `to=png/pdf/svg` uniform; `to=md` output form differs (docx/xlsx→markdown, pptx→per-slide SVG) |
| 4 | `set_doc_text` | EDIT (det) | no | ✅ | ✅ | ✅ | **action vocab differs** (see below) |
| 5 | `arrange_doc` | EDIT (det) | no | ❌ | ✅ | ✅ | **docx rejected**; ops differ (pptx: duplicate/move/delete slide · xlsx: +rename sheet) |
| 6 | `read_doc_xml` | EDIT (det) | no | ✅ | ✅ | ✅ | format-agnostic (any OOXML zip of XML parts) |
| 7 | `set_doc_xml` | EDIT (det) | no | ✅ | ✅ | ✅ | format-agnostic |
| 8 | `build_doc` | GENERATE (det) | no | ✅ | ✅ | ✅ | **spec shape differs**: docx←markdown · xlsx←`{sheets}` · pptx←`{slides,theme?}` |
| 9 | `generate_doc` | GENERATE (LLM) | yes | ✅ | ✅ | ✅ | pptx = full deck pipeline; some params PPTX-only |
| — | `edit_doc` | EDIT (LLM) | yes | ✅ | ✅ | ✅ | pptx rerouted to the deck editor; docx/xlsx via `edit_doc.py` |

**Per-format op vocabulary that lives *inside* a verb** (the real format-specific surface):

| verb | .docx | .xlsx | .pptx |
|---|---|---|---|
| `set_doc_text` edits | `replace` / `insert_after` / `delete` (para or table cell) | `set_cell` / `append_rows` / `add_sheet` | paragraph text at `{slide,shape_id,para}` + `{chart,…}` routing |
| `arrange_doc` ops | *(none — rejected)* | `duplicate`(needs name) / `move` / `delete` / `rename` | `duplicate` / `move` / `delete` (slide index) |
| `build_doc` spec | markdown string | `{"sheets":[…]}` | `{"slides":[…], "theme":{…}?}` |
| `edit_doc` actions | `replace` / `insert_after` / `delete` | `set_cell` / `append_rows` / `add_sheet` | *(deck pipeline — separate vocab)* |

### Two more tiers (for completeness — not the agent surface)

- **Library-only functions** (`import edit2docs`): `edit_chart`, `list_charts`,
  `list_doc_parts`, `get_doc_xml`, `preview_doc`, plus a **PPTX-only** convenience
  set (`generate_pptx`, `edit_pptx`, `preview_pptx`, `set_pptx_text`, `analyze_pptx`)
  and `async_*` variants. These back the 9 verbs; agents don't call them directly.
- **geny-executor `Doc*` built-ins** — 1:1 remap of the 9 verbs (+`DocArrange`):
  `DocGuide/DocAnalyze/DocRender/DocApplyEdits/DocArrange/DocXmlRead/DocXmlEdit/DocBuild/DocGenerate/DocEdit`.
- **Hosted MCP `mcp/server.py`** — a separate, asset-oriented, **PPTX-centric** set
  (`generate_deck`, `edit_deck`, `upload_source`, `get_asset`, …), not the 9 verbs.

---

## Part 2 — Current exposure mechanism (why there is no extension hierarchy)

- The tool list is **static and format-blind**. `ANTHROPIC_TOOLS` is a hand-written
  list of 9 dicts; `OPENAI_TOOLS`/`TOOL_NAMES` derive from it; local MCP registers
  all 10 decorators at server start. **No surface takes a file/extension/format
  parameter at registration time.**
- `tool_specs(fmt)` — the `fmt` is **`anthropic`|`openai`** (a *provider* selector),
  **not** a document format. There is no `tool_specs("xlsx")`.
- Format is discovered **per call**, from the `doc`/`output` path suffix
  (`simple._fmt_of`). So the engine already knows the format at execution time — but
  the *agent* was shown every tool and every format's vocabulary up front.

**What an agent editing `report.xlsx` needlessly sees today:**
- `arrange_doc` (which will only ever return `invalid` for… no — xlsx is fine; but a
  `.docx` user sees `arrange_doc`, which always rejects docx).
- `set_doc_text`'s description packing docx+xlsx+pptx actions together.
- `doc_guide` ROOT advertising `recipes.slides` and `arrange` to a docx user.
- `generate_doc`'s "PPTX only" params.

---

## Part 3 — What "hierarchy by extension" concretely means

Honest framing: because edit2docs **deliberately unified** its verbs across formats
(the v0.12 consolidation to 8), extension scoping at the **verb level removes almost
nothing** — only `arrange_doc` is truly format-restricted (drop it for `.docx`).
The real hierarchy operates on **three levels**, in increasing value:

1. **Verb selection** (thin) — hide verbs that don't apply: `arrange_doc`→ hidden for
   `.docx`. (If we later add docx-flow-arrange, even this disappears.)
2. **Schema / description specialization** (the main win) — each verb's *description*
   and *input schema* show **only the target format's vocabulary**. `set_doc_text` for
   `.xlsx` advertises `set_cell/append_rows/add_sheet` and nothing about slides or
   charts; `arrange_doc` for `.pptx` constrains `op` to slide ops. The agent stops
   trying inapplicable operations and reads a shorter, sharper schema.
3. **Guide scoping** (`doc_guide`) — the map and topic list are filtered: a `.docx`
   session's `doc_guide()` omits `arrange`, `recipes.slides`, the pptx half of
   `edit.chart`, etc.

The token saving is modest (descriptions are ≤320 chars); the **correctness and
focus** gain is the point — the agent is guided to the right operation the first time.

---

## Part 4 — Proposed design

### 4.1 A format-capability table (single source of truth)

One module-level table in edit2docs (`agent_tools` or a new `tool_matrix.py`) maps
`(verb, format)` → `{applicable: bool, description: str, schema_fragment: dict}`.
Every scoped surface (schemas, guide, MCP) derives from it — no drift.

```python
CAPABILITY = {
    "arrange_doc": {
        "docx": None,  # not applicable → hidden
        "xlsx": {"desc": "... sheets: duplicate/move/rename/delete ...",
                 "ops_enum": ["duplicate", "move", "delete", "rename"]},
        "pptx": {"desc": "... slides: duplicate/move/delete ...",
                 "ops_enum": ["duplicate", "move", "delete"]},
    },
    "set_doc_text": { "docx": {...}, "xlsx": {...}, "pptx": {...} },
    ...  # verbs applicable to all 3 with identical schema list only their desc
}
```

### 4.2 Extension-aware `tool_specs`

Extend (keep back-compat): `tool_specs(provider="anthropic", *, extension=None)`.
- `extension=None` → today's full, generic 9-verb list (unchanged default).
- `extension="xlsx"` → only applicable verbs, each with its **xlsx-specialized**
  description + schema (from the capability table). `arrange_doc` present; a `.docx`
  call would omit it.
- `extension` may be a **set** (`{"xlsx","pptx"}`) for multi-file sessions → union of
  applicable verbs, each description noting which formats it targets.

(Rename the existing `fmt` param to `provider` with a deprecation alias, since `fmt`
now reads ambiguously next to a document extension.)

### 4.3 Extension-aware `doc_guide`

`doc_guide(topic=None, *, fmt=None)`:
- `fmt=None` → today's full map (unchanged).
- `fmt="docx"` → ROOT map lists only all-format + docx topics; `TOPICS` excludes
  `arrange`, `recipes.slides`; `edit.text`/`edit.chart`/`build` render only the docx
  block. The per-topic guide text is sliced from the same capability table.

### 4.4 Multi / none / unknown policy

| situation | exposed |
|---|---|
| one file, known ext | scoped+specialized set for that ext |
| several files, mixed ext | union of applicable verbs; per-verb desc lists the relevant formats |
| no file yet (pure generate) | GENERATE family (`build_doc`, `generate_doc`) + `doc_guide` only |
| unknown / no extension | full generic set (today's behavior) — never *less* capable than now |

---

## Part 5 — Consumer adoption (where the scoping actually happens)

The engine knows the format per call; the **agent-facing tool list** is chosen by the
*host*. Adoption differs by host because of how each registers tools:

- **Library** (`import edit2docs`) — unchanged; all functions always available.
- **geny-executor (primary win).** geny **knows the attached file(s)** in the session
  working dir and already re-emits its tool set within a turn (`tools/list_changed`,
  progressive disclosure). When a document is attached, geny registers only the
  scoped `Doc*` set with format-specialized descriptions (built from
  `edit2docs.tool_specs(extension=…)`). This is where irrelevant-tool-noise is
  actually removed for the model. **No file → the model calls `DocGuide(file)` /
  attaches one, then the set re-scopes.**
- **Local stdio MCP (`edit2docs-mcp`).** MCP registers tools **once at startup**,
  *before* any file is known — so it cannot pre-scope by extension. Three options:
  (a) keep the full 10 tools but make **`doc_guide` the extension-aware entry point**
  (the model passes the file; the guide scopes itself) — zero-infra, recommended
  default; (b) honor an `EDIT2DOCS_FORMAT` launch env to scope a single-format
  server; (c) emit `tools/list_changed` after the first call reveals the working
  file (dynamic re-scoping) for clients that support it.
- **Hosted MCP `mcp/server.py` / `/v1/text-edits`** — already PPTX-scoped in practice;
  optionally generalize later, out of scope here.

---

## Part 6 — Design tension (needs your call)

- **A. Scope + specialize the unified verbs** *(recommended)* — keep the 9 verbs;
  filter the set and tailor each verb's description/schema/guide by extension via the
  capability table. Preserves the consolidation philosophy; small, additive.
- **B. Extension-namespaced tools** — expose format-specific tool *names*
  (`xlsx_set_cell`, `pptx_arrange_slides`, …). Maximally "only relevant tools", but
  **re-expands** the surface the v0.12 consolidation deliberately shrank (more tools,
  more names, heavier guide). Not recommended.
- **C. Hybrid** — unified verbs stay; add a thin `tools_for(extension)` manifest layer
  hosts call. This is A's mechanism; list it only to name the seam.

Recommendation: **A** (mechanism = C's `tool_specs(extension=…)` + capability table).

---

## Part 7 — Decisions for your review

- **D1. Scope granularity** — A (scope+specialize unified verbs, *recommended*) vs.
  B (extension-namespaced tool names).
- **D2. Default when extension unknown** — full generic set (*recommended*, never
  regress) vs. force the model to declare a format first.
- **D3. Local-MCP strategy** — `doc_guide`-as-entry-point (*recommended*, zero-infra)
  vs. `EDIT2DOCS_FORMAT` launch env vs. `tools/list_changed` dynamic re-scoping.
- **D4. Where the capability table lives** — new `tool_matrix.py` (SSOT) that
  `agent_tools` + `agent_guide` both import (*recommended*) vs. inline in `agent_tools`.
- **D5. geny exposure** — geny scopes `Doc*` by the attached file's extension
  (*recommended*; the real payoff) — confirm this is in scope for the same effort.
- **D6. `tool_specs` signature** — add `extension=` and rename `fmt`→`provider` with a
  back-compat alias (*recommended*) vs. a separate `tool_specs_for_ext()`.

---

## Part 8 — Milestones (once decisions land)

1. **M1** — `tool_matrix.py` capability table (SSOT) + unit tests asserting every
   `(verb, format)` cell and that applicable-verb sets match the engines.
2. **M2** — `tool_specs(provider, *, extension=)` scoping + format-specialized
   schemas/descriptions; `doc_guide(topic, *, fmt=)` scoping; update the three
   hard-coded surface tests. → edit2docs minor release.
3. **M3** — geny-executor: scope the `Doc*` registration by the session's attached
   document extension (build descriptions from `tool_specs(extension=…)`). → release.
4. **M4** — deploy hr-web + align geny (as usual). Local MCP: `doc_guide` entry-point
   scoping ships with M2; `EDIT2DOCS_FORMAT` / `list_changed` optional follow-ups.

Estimated: mostly declarative (the capability table) + wiring; low risk, no engine
behavior change — only *which tool metadata the agent is shown*.
