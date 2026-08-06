"""Format-capability matrix — the single source of truth for which agent
verb applies to which document format, and the format-specialized blurb each
verb shows once the extension is known.

Consumed by ``agent_tools.tool_specs(extension=...)`` and
``agent_guide.doc_guide(fmt=...)``. Both are **opt-in**: with no extension the
full, generic tool set / topic list is returned exactly as before — scoping
never makes a caller *less* capable than the unscoped default.

Why this exists: the verbs were deliberately unified across formats (v0.12),
so extension scoping removes little at the *verb* level — only ``arrange_doc``
is truly format-restricted. The real win is (1) hiding that one verb for
``.docx``, (2) showing each verb's description tuned to the target format
(``set_doc_text`` for ``.xlsx`` advertises cells, not slides), and (3)
trimming the ``doc_guide`` topic list. All three read from this table.
"""

from __future__ import annotations

from typing import Any

FORMATS: tuple[str, ...] = ("docx", "xlsx", "pptx")

#: Which formats each agent verb applies to. Everything is all-format except
#: arrange_doc — docx is a flow document with no slide/sheet structure.
VERB_FORMATS: dict[str, set[str]] = {
    "doc_guide": set(FORMATS),
    "analyze_doc": set(FORMATS),
    "render_doc": set(FORMATS),
    "set_doc_text": set(FORMATS),
    "arrange_doc": {"xlsx", "pptx"},
    "read_doc_xml": set(FORMATS),
    "set_doc_xml": set(FORMATS),
    "build_doc": set(FORMATS),
    "generate_doc": set(FORMATS),
    "edit_doc": set(FORMATS),
}

#: Per-(verb, format) description override (≤320 chars, like the base set).
#: Only verbs whose op vocabulary or spec shape is format-specific appear
#: here; anything absent keeps its generic ANTHROPIC_TOOLS description.
DESCRIPTION_OVERRIDES: dict[str, dict[str, str]] = {
    "set_doc_text": {
        "docx": (
            "Deterministic paragraph / table-cell edits at analyze_doc "
            "addresses — actions replace | insert_after | delete. No key; "
            "byte-preserves the rest. Shapes: doc_guide('edit.text')."
        ),
        "xlsx": (
            "Deterministic cell edits at analyze_doc addresses — actions "
            "set_cell | append_rows | add_sheet. No key; byte-preserves the "
            "rest. Shapes: doc_guide('edit.text')."
        ),
        "pptx": (
            "Deterministic text edits at analyze_doc addresses (slide / shape "
            "/ para) AND chart title/data ({chart:i,...}). No key; byte-"
            "preserves the rest. Shapes: doc_guide('edit.text'), "
            "doc_guide('edit.chart')."
        ),
    },
    "arrange_doc": {
        "xlsx": (
            "Deterministic STRUCTURAL edits: duplicate / move / delete / "
            "rename whole sheets. No key; byte-preserving. ops apply in order "
            "— target = sheet name or index, to = position. Guide: "
            "doc_guide('arrange')."
        ),
        "pptx": (
            "Deterministic STRUCTURAL edits: duplicate / move / delete whole "
            "slides. No key; byte-preserving. ops apply in order — target = "
            "slide index, to = position. Guide: doc_guide('arrange')."
        ),
    },
    "build_doc": {
        "docx": (
            "GENERATE (deterministic): a NEW .docx from a MARKDOWN spec "
            "string. No LLM, no key. Options: doc_guide('build')."
        ),
        "xlsx": (
            "GENERATE (deterministic): a NEW .xlsx from a {\"sheets\":[...]} "
            "spec. No LLM, no key. Options: doc_guide('build')."
        ),
        "pptx": (
            "GENERATE (deterministic): a NEW .pptx from a {\"slides\":[...], "
            "\"theme\"?} spec — themed decks in one call. No LLM. Options: "
            "doc_guide('build')."
        ),
    },
    "generate_doc": {
        "pptx": (
            "GENERATE (LLM): a complete designed deck from a one-line intent "
            "(Anthropic key; .pptx is slow — minutes). Options: "
            "doc_guide('generate')."
        ),
    },
}

#: Which formats each doc_guide topic is relevant to (used by
#: doc_guide(fmt=...) to trim the topic list).
TOPIC_FORMATS: dict[str, set[str]] = {
    "build": set(FORMATS),
    "generate": set(FORMATS),
    "edit": set(FORMATS),
    "edit.text": set(FORMATS),
    "edit.chart": {"xlsx", "pptx"},
    "arrange": {"xlsx", "pptx"},
    "edit.xml": set(FORMATS),
    "recipes.slides": {"pptx"},
    "recipes.colors": {"pptx"},
    "render": set(FORMATS),
}


def normalize_extension(extension: Any) -> set[str] | None:
    """A validated set of document formats, or None (= no scoping).

    Accepts a single extension (``"xlsx"`` / ``".xlsx"`` / a full path),
    or an iterable of them. Unknown / empty inputs → None (fall back to the
    full unscoped set), so scoping never regresses to *fewer* than today when
    the caller isn't sure of the format.
    """
    if extension is None:
        return None
    items = [extension] if isinstance(extension, str) else list(extension)
    out: set[str] = set()
    for item in items:
        ext = str(item).lower().rsplit(".", 1)[-1].strip()
        if ext in FORMATS:
            out.add(ext)
    return out or None


def topics_for(fmt: Any) -> list[str] | None:
    """The doc_guide topic list scoped to *fmt*, or None if unscoped."""
    exts = normalize_extension(fmt)
    if exts is None:
        return None
    return [t for t, fs in TOPIC_FORMATS.items() if not fs.isdisjoint(exts)]


def scope_tools(base_tools: list[dict[str, Any]], extension: Any) -> list[dict[str, Any]]:
    """Filter + format-specialize *base_tools* for the given extension(s).

    * verbs not applicable to any provided format are dropped
      (``arrange_doc`` for a ``.docx`` session);
    * with a single known format, each verb's description is swapped for its
      format-specialized variant where one exists;
    * with several formats (mixed-file session) or none, descriptions stay
      generic.

    Returns a NEW list of NEW dicts — the module-level ``ANTHROPIC_TOOLS`` is
    never mutated.
    """
    exts = normalize_extension(extension)
    if exts is None:
        return [dict(t) for t in base_tools]
    single = next(iter(exts)) if len(exts) == 1 else None
    out: list[dict[str, Any]] = []
    for tool in base_tools:
        applies = VERB_FORMATS.get(tool["name"], set(FORMATS))
        if applies.isdisjoint(exts):
            continue
        spec = dict(tool)
        if single is not None:
            override = DESCRIPTION_OVERRIDES.get(tool["name"], {}).get(single)
            if override is not None:
                spec["description"] = override
        out.append(spec)
    return out


def applicable_verbs(extension: Any) -> list[str] | None:
    """Verb names applicable to *extension* (order follows the base list),
    or None if unscoped. Handy for hosts that scope their own tool registry
    (e.g. geny-executor's Doc* built-ins)."""
    exts = normalize_extension(extension)
    if exts is None:
        return None
    return [
        name
        for name, fs in VERB_FORMATS.items()
        if not fs.isdisjoint(exts)
    ]


def describe(verb: str, extension: Any) -> str | None:
    """The format-specialized description for *verb* at a single *extension*,
    or None to signal 'use the generic one'. Exposed for host tool layers
    (geny DocGuide/DocApplyEdits…) that want the same specialized blurbs."""
    exts = normalize_extension(extension)
    if exts is None or len(exts) != 1:
        return None
    (fmt,) = tuple(exts)
    return DESCRIPTION_OVERRIDES.get(verb, {}).get(fmt)


__all__ = [
    "FORMATS",
    "VERB_FORMATS",
    "TOPIC_FORMATS",
    "DESCRIPTION_OVERRIDES",
    "normalize_extension",
    "topics_for",
    "scope_tools",
    "applicable_verbs",
    "describe",
]
