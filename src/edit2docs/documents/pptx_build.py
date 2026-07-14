"""Deterministic .pptx construction from a slide spec (no LLM, no key).

The sibling of :func:`docx_from_markdown` and :func:`xlsx_from_spec`: given a
structured slide list, render a standard, fully-editable PowerPoint deck using
python-pptx's built-in slide layouts. There is **no layout intelligence** here
— that is precisely what the LLM ``generate_pptx`` pipeline provides. This
builder is the deterministic primitive an agent (or that pipeline) drives when
it already knows what each slide should say.

Spec shape::

    {"slides": [
        {"layout": "title",   "title": "Deck Title", "subtitle": "Q3 2026"},
        {"layout": "section", "title": "Part One"},
        {"layout": "content", "title": "Agenda",
         "bullets": ["First point", {"text": "Sub point", "level": 1}]},
        {"layout": "title_only", "title": "Just a heading"},
        {"layout": "blank"},
    ]}

Each slide may also carry ``"notes"`` (speaker notes, plain text). ``layout``
is optional and defaults to ``content`` (title + bullets). Unknown layout
names fall back to ``content``. An empty / missing ``slides`` list raises
``ValueError`` (bilingual), mirroring :func:`xlsx_from_spec`.
"""

from __future__ import annotations

from typing import Any

__all__ = ["pptx_from_spec"]

# Canonical layout name -> python-pptx default-template slide_layouts index.
# The default template ships these nine layouts in a stable order.
_LAYOUT_INDEX = {
    "title": 0,          # Title Slide (title + subtitle)
    "title_content": 1,  # Title and Content (title + body placeholder)
    "content": 1,        # alias
    "bullets": 1,        # alias
    "section": 2,        # Section Header
    "two_content": 3,    # Two Content
    "comparison": 4,     # Comparison
    "title_only": 5,     # Title Only
    "blank": 6,          # Blank
}
_DEFAULT_LAYOUT = "content"


def _coerce_bullets(raw: Any) -> list[tuple[str, int]]:
    """Normalize a bullets value into ``[(text, level), ...]``.

    Accepts a list of strings or ``{"text": str, "level": int}`` dicts. A
    plain string is treated as a single bullet. Levels are clamped to 0..8.
    """
    if raw is None:
        return []
    if isinstance(raw, str):
        raw = [raw]
    out: list[tuple[str, int]] = []
    for item in raw:
        if isinstance(item, dict):
            text = str(item.get("text", ""))
            try:
                level = int(item.get("level", 0) or 0)
            except (TypeError, ValueError):
                level = 0
        else:
            text, level = str(item), 0
        out.append((text, max(0, min(8, level))))
    return out


def _set_title(slide, text: str) -> None:
    """Set the slide's title placeholder, if it has one."""
    if not text:
        return
    title_ph = slide.shapes.title
    if title_ph is not None:
        title_ph.text = str(text)


def _first_body_placeholder(slide):
    """The first non-title placeholder that can hold body text, or None.

    Placeholder idx 0 is the title. We must compare by ``idx``, NOT by
    identity: ``slide.shapes.title`` and the placeholder yielded while
    iterating ``slide.placeholders`` are distinct proxy objects wrapping the
    same element, so ``ph is title`` is always False and would let us clobber
    the title. SUBTITLE (4) / BODY (2) / OBJECT (7) / CONTENT all work as the
    body target — anything with a text frame that isn't the title.
    """
    for ph in slide.placeholders:
        if ph.placeholder_format is not None and ph.placeholder_format.idx == 0:
            continue  # the title placeholder
        if ph.has_text_frame:
            return ph
    return None


def _fill_bullets(slide, bullets: list[tuple[str, int]]) -> None:
    if not bullets:
        return
    body = _first_body_placeholder(slide)
    if body is None:
        return
    tf = body.text_frame
    for i, (text, level) in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = text
        para.level = level


def _set_subtitle(slide, text: str) -> None:
    """Fill the subtitle/body placeholder on a title-style slide."""
    if not text:
        return
    body = _first_body_placeholder(slide)
    if body is not None:
        body.text_frame.text = str(text)


def _set_notes(slide, text: str) -> None:
    if not text:
        return
    slide.notes_slide.notes_text_frame.text = str(text)


def pptx_from_spec(spec: dict[str, Any], *, lang: str | None = None) -> bytes:
    """Render a slide spec into a .pptx package. Deterministic, no LLM.

    Raises ``ValueError`` on a structurally invalid spec (bilingual message)
    — the ``generate`` retry loop, or an agent, can react to that signal.
    """
    from io import BytesIO

    from pptx import Presentation

    slides = spec.get("slides") if isinstance(spec, dict) else None
    if not isinstance(slides, list) or not slides:
        raise ValueError(
            "slide spec must contain a non-empty `slides` list. "
            "슬라이드 스펙에는 비어있지 않은 `slides` 목록이 필요합니다."
        )

    prs = Presentation()
    for i, raw in enumerate(slides):
        if not isinstance(raw, dict):
            raise ValueError(
                f"slide {i} must be an object. "
                f"슬라이드 {i}는 객체여야 합니다."
            )
        name = str(raw.get("layout") or _DEFAULT_LAYOUT).strip().lower()
        idx = _LAYOUT_INDEX.get(name, _LAYOUT_INDEX[_DEFAULT_LAYOUT])
        slide = prs.slides.add_slide(prs.slide_layouts[idx])

        _set_title(slide, raw.get("title", ""))
        if name == "title":
            _set_subtitle(slide, raw.get("subtitle", ""))
        elif idx == 1 or "bullets" in raw:
            _fill_bullets(slide, _coerce_bullets(raw.get("bullets")))
        _set_notes(slide, raw.get("notes", ""))

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()
