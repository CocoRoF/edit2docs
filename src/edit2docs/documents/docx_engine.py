"""Deterministic DOCX building blocks (no LLM).

Markdown is the interchange format: the generator LLM writes markdown and
``docx_from_markdown`` renders it; ``docx_to_markdown`` goes the other way
(via mammoth) so editors can read a document back. Paragraph-level
addressing (``docx_outline`` / ``apply_docx_edits``) keeps untouched
paragraphs byte-identical — same philosophy as the PPTX recompose path.

Supported markdown subset (generation):
  # .. ###### headings · paragraphs · - / * bullets · 1. numbered lists ·
  **bold** / *italic* / `code` inline · pipe tables · > blockquote ·
  --- horizontal rule · ``` fenced code blocks (monospace paragraphs)
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

__all__ = [
    "docx_from_markdown",
    "docx_to_markdown",
    "docx_to_html",
    "docx_outline",
    "apply_docx_edits",
    "DocxEdit",
    "DocxEditResult",
]


# ---------------------------------------------------------------------------
# Markdown -> DOCX
# ---------------------------------------------------------------------------

_INLINE = re.compile(
    r"(\*\*(?P<bold>.+?)\*\*)|(\*(?P<italic>[^*]+?)\*)|(`(?P<code>[^`]+?)`)"
)


def _add_runs(paragraph, text: str) -> None:
    """Render **bold** / *italic* / `code` inline markup into runs."""
    pos = 0
    for match in _INLINE.finditer(text):
        if match.start() > pos:
            paragraph.add_run(text[pos : match.start()])
        if match.group("bold") is not None:
            paragraph.add_run(match.group("bold")).bold = True
        elif match.group("italic") is not None:
            paragraph.add_run(match.group("italic")).italic = True
        else:
            run = paragraph.add_run(match.group("code"))
            run.font.name = "D2Coding"
        pos = match.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def _is_table_row(line: str) -> bool:
    stripped = line.strip()
    return stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 2


def _split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _is_separator_row(line: str) -> bool:
    """True only for genuine markdown alignment rows (---, :---:, ...).

    A permissive character-class match also swallowed real data rows like
    ``| - | - |`` — require every cell to be ``:?-{2,}:?``.
    """
    cells = _split_table_row(line)
    return bool(cells) and all(
        re.fullmatch(r":?-{2,}:?", cell) for cell in cells
    )


def docx_from_markdown(markdown: str, *, base_font: str = "맑은 고딕") -> bytes:
    """Render the supported markdown subset into a .docx package."""
    document = Document()
    style = document.styles["Normal"]
    style.font.name = base_font
    style.font.size = Pt(10.5)
    # East-Asian typeface must be set on the rPr explicitly.
    style.element.rPr.rFonts.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia",
        base_font,
    )

    lines = markdown.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # Fenced code block -> monospace paragraphs.
        if stripped.startswith("```"):
            i += 1
            code_lines: list[str] = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1  # closing fence
            for code_line in code_lines or [""]:
                p = document.add_paragraph()
                run = p.add_run(code_line)
                run.font.name = "D2Coding"
                run.font.size = Pt(9)
            continue

        # Pipe table.
        if _is_table_row(stripped):
            rows: list[list[str]] = []
            while i < len(lines) and _is_table_row(lines[i].strip()):
                if not _is_separator_row(lines[i]):
                    rows.append(_split_table_row(lines[i]))
                i += 1
            if rows:
                cols = max(len(r) for r in rows)
                table = document.add_table(rows=len(rows), cols=cols)
                table.style = "Table Grid"
                for r_idx, row in enumerate(rows):
                    for c_idx in range(cols):
                        cell_text = row[c_idx] if c_idx < len(row) else ""
                        cell_p = table.rows[r_idx].cells[c_idx].paragraphs[0]
                        _add_runs(cell_p, cell_text)
                        if r_idx == 0:
                            for run in cell_p.runs:
                                run.bold = True
            continue

        # Heading.
        heading = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if heading:
            level = len(heading.group(1))
            p = document.add_heading("", level=min(level, 6))
            _add_runs(p, heading.group(2))
            i += 1
            continue

        # Horizontal rule -> centered separator.
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
            p = document.add_paragraph("⸻")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            i += 1
            continue

        # Blockquote.
        if stripped.startswith(">"):
            p = document.add_paragraph(style="Intense Quote")
            _add_runs(p, stripped.lstrip("> ").strip())
            i += 1
            continue

        # Bullet / numbered list item.
        bullet = re.match(r"^[-*]\s+(.*)$", stripped)
        if bullet:
            p = document.add_paragraph(style="List Bullet")
            _add_runs(p, bullet.group(1))
            i += 1
            continue
        numbered = re.match(r"^\d+[.)]\s+(.*)$", stripped)
        if numbered:
            p = document.add_paragraph(style="List Number")
            _add_runs(p, numbered.group(1))
            i += 1
            continue

        # Plain paragraph (consume soft-wrapped continuation lines).
        chunk = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt.startswith(("#", ">", "```", "-", "*"))
                or _is_table_row(nxt)
                or re.match(r"^\d+[.)]\s", nxt)
            ):
                break
            chunk.append(nxt)
            i += 1
        p = document.add_paragraph()
        _add_runs(p, " ".join(chunk))

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# DOCX -> markdown / outline
# ---------------------------------------------------------------------------


def docx_to_markdown(content: bytes) -> str:
    """Convert a .docx to markdown via mammoth (structure-preserving)."""
    import mammoth
    from markdownify import markdownify

    html = mammoth.convert_to_html(io.BytesIO(content)).value
    return markdownify(html, heading_style="ATX").strip()


_HEADING_STYLE = re.compile(r"[Hh]eading\s*(\d)")


def docx_outline(content: bytes) -> list[dict]:
    """Paragraph outline with the addresses ``apply_docx_edits`` accepts.

    Entries: ``{"para": i, "style", "text"}`` for body paragraphs and
    ``{"table": t, "row": r, "col": c, "text"}`` for table cells (first
    paragraph of each cell).
    """
    document = Document(io.BytesIO(content))
    outline: list[dict] = []
    for i, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        style = paragraph.style.name if paragraph.style is not None else "Normal"
        outline.append({"para": i, "style": style, "text": text})
    for t, table in enumerate(document.tables):
        for r, row in enumerate(table.rows):
            for c, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text:
                    outline.append({"table": t, "row": r, "col": c, "text": text})
    return outline


# ---------------------------------------------------------------------------
# Targeted edits
# ---------------------------------------------------------------------------


@dataclass
class DocxEdit:
    """One operation. ``action``:

    * ``replace`` — replace paragraph ``para``'s text (formatting of the
      first run survives) or table cell (``table``+``row``+``col``).
    * ``insert_after`` — insert markdown-rendered paragraphs after ``para``
      (para=-1 inserts at the start).
    * ``delete`` — remove paragraph ``para``.
    """

    action: str
    para: int | None = None
    table: int | None = None
    row: int | None = None
    col: int | None = None
    new_text: str = ""
    old_text: str | None = None
    markdown: str = ""


@dataclass
class DocxEditResult:
    action: str
    status: str  # applied | stale | not_found | invalid
    message: str = ""


def _normalize(text: str) -> str:
    return " ".join(text.split())


def apply_docx_edits(content: bytes, edits: Iterable[DocxEdit]) -> tuple[bytes, list[DocxEditResult]]:
    """Apply edits; untouched paragraphs stay byte-identical.

    Per-edit soft failures (like the PPTX text editor): ``old_text``
    guards replaces with a whitespace-normalized comparison.
    """
    document = Document(io.BytesIO(content))
    edit_list = list(edits)
    results: list[DocxEditResult | None] = [None] * len(edit_list)

    # Deletions/insertions shift indices — apply ops sorted by paragraph
    # DESCENDING so every edit's original address stays valid — but report
    # results in the CALLER'S order (clients correlate results[i] <-> edits[i]).
    ordered = sorted(
        enumerate(edit_list),
        key=lambda pair: (pair[1].para if pair[1].para is not None else 10**9),
        reverse=True,
    )
    for index, edit in ordered:
        results[index] = _apply_one(document, edit)

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue(), [r for r in results if r is not None]


def _apply_one(document, edit: DocxEdit) -> DocxEditResult:
    if edit.action == "replace" and edit.table is not None:
        if edit.row is None or edit.col is None:
            return DocxEditResult(edit.action, "invalid", "table replace needs row+col")
        try:
            cell = document.tables[edit.table].rows[edit.row].cells[edit.col]
        except IndexError:
            return DocxEditResult(edit.action, "not_found", "table cell out of range")
        if edit.old_text is not None and _normalize(cell.text) != _normalize(edit.old_text):
            return DocxEditResult(edit.action, "stale", "cell text changed; refresh")
        _set_paragraph_text(cell.paragraphs[0], edit.new_text)
        for extra in cell.paragraphs[1:]:
            _set_paragraph_text(extra, "")
        return DocxEditResult(edit.action, "applied")

    para_ok = (
        edit.para is not None
        and 0 <= edit.para < len(document.paragraphs)
    ) or (edit.action == "insert_after" and edit.para == -1)
    if not para_ok:
        return DocxEditResult(edit.action, "not_found", "paragraph index out of range")

    if edit.action == "replace":
        paragraph = document.paragraphs[edit.para]
        if edit.old_text is not None and _normalize(paragraph.text) != _normalize(edit.old_text):
            return DocxEditResult(edit.action, "stale", "paragraph text changed; refresh")
        _set_paragraph_text(paragraph, edit.new_text)
        return DocxEditResult(edit.action, "applied")

    if edit.action == "delete":
        paragraph = document.paragraphs[edit.para]
        p = paragraph._p
        p.getparent().remove(p)
        return DocxEditResult(edit.action, "applied")

    if edit.action == "insert_after":
        rendered = Document(io.BytesIO(docx_from_markdown(edit.markdown or edit.new_text)))
        # Harvest block elements (paragraphs AND tables) in document order;
        # skip empty paragraphs and the section properties element.
        from docx.oxml.ns import qn as _qn

        new_elements = []
        for element in rendered.element.body:
            if element.tag == _qn("w:sectPr"):
                continue
            if element.tag == _qn("w:p") and not "".join(element.itertext()).strip():
                continue
            new_elements.append(element)
        if not new_elements:
            return DocxEditResult(edit.action, "invalid", "nothing to insert")
        if edit.para == -1:
            anchor = document.paragraphs[0]._p if document.paragraphs else None
            for element in new_elements:
                if anchor is None:
                    document.element.body.append(element)
                else:
                    anchor.addprevious(element)
        else:
            anchor = document.paragraphs[edit.para]._p
            for element in reversed(new_elements):
                anchor.addnext(element)
        return DocxEditResult(edit.action, "applied")

    return DocxEditResult(edit.action, "invalid", f"unknown action {edit.action!r}")


def _set_paragraph_text(paragraph, new_text: str) -> None:
    """First run keeps its formatting; the rest are emptied.

    Hyperlink children (``w:hyperlink``) hold their own runs that
    ``paragraph.runs`` doesn't expose — remove them so the old link text
    doesn't survive next to the replacement.
    """
    from docx.oxml.ns import qn

    for hyperlink in paragraph._p.findall(qn("w:hyperlink")):
        paragraph._p.remove(hyperlink)
    text = " ".join(new_text.split("\n"))
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


_SAFE_HREF = re.compile(r"^(https?:|mailto:)", re.IGNORECASE)
_ANCHOR_HREF = re.compile(r'<a\s+href="([^"]*)"')


def _sanitize_preview_html(html: str) -> str:
    """Neutralize unsafe URL schemes in anchors (javascript:, data:, ...).

    mammoth emits only structural tags, but hyperlink hrefs come verbatim
    from the docx relationships — a crafted document can carry
    ``javascript:`` URLs. Allowlist http/https/mailto; safe links open in
    a new tab so the studio SPA keeps its state.
    """

    def _fix(match: re.Match) -> str:
        href = match.group(1)
        if not _SAFE_HREF.match(href):
            return "<a"  # drop the href entirely
        return f'<a target="_blank" rel="noopener noreferrer" href="{href}"'

    return _ANCHOR_HREF.sub(_fix, html)


def docx_to_html(content: bytes) -> str:
    """Convert a .docx to display HTML via mammoth (headings, lists, tables).

    Anchor hrefs are sanitized to http/https/mailto — see
    :func:`_sanitize_preview_html` — so the result is safe to inline in a
    styled preview container.
    """
    import mammoth

    return _sanitize_preview_html(mammoth.convert_to_html(io.BytesIO(content)).value)
