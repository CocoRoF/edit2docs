"""Recompose a PPTX: keep / replace / insert / delete / reorder slides.

The chat-editing pipeline (``tools/edit_deck``) reduces every edit to one
primitive: the final deck is an ordered *sequence* whose entries are either
an original slide (kept as-is, identity preserved) or a freshly generated
SVG slide. That single primitive covers:

* replace slide N   -> sequence swaps ``KeepSlide(N)`` for a ``NewSlide``
* insert after N    -> ``NewSlide`` spliced into the sequence
* delete slide N    -> ``KeepSlide(N)`` simply absent
* reorder           -> permuted ``KeepSlide`` entries

Original slides not present in the sequence are physically removed from
the package (parts, rels, notesSlides, content-type overrides). Kept
slides keep their part names, slide ids, notes and animations untouched.

Same raw-OOXML style as ``pptx_append`` (string/regex surgery on the
extracted package), and it reuses that module's helpers.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

from .pptx_append import (
    AppendError,
    SLIDE_CONTENT_TYPE,
    SLIDE_REL_TYPE,
    _add_default_content_type,
    _add_override,
    _existing_slides,
    _next_free_sld_id,
    _next_free_slide_number,
    _pick_blankest_layout,
    remove_slide_parts,
    write_slide_part_from_svg,
)
from .pptx_builder import _append_relationship, _content_type_for_extension


@dataclass(frozen=True)
class KeepSlide:
    """An original slide, addressed by its 0-based position in the current deck."""

    index: int


@dataclass(frozen=True)
class NewSlide:
    """A freshly generated slide rendered from *svg_path*."""

    svg_path: Path


def recompose_pptx(
    host_pptx: Path,
    sequence: list[KeepSlide | NewSlide],
    output_path: Path,
    *,
    transition: str | None = "fade",
    transition_duration: float = 0.5,
    verbose: bool = False,
) -> list[dict]:
    """Rebuild *host_pptx* so its slides are exactly *sequence*, in order.

    Returns non-fatal warnings as ``{"code", "message"}`` dicts.

    Raises:
        ValueError: empty sequence, out-of-range or duplicated KeepSlide index.
        AppendError: the host package is not a usable PPTX.
    """
    if not sequence:
        raise ValueError("recompose_pptx requires a non-empty slide sequence")

    warnings: list[dict] = []
    temp_dir = Path(tempfile.mkdtemp(prefix="edit2docs-recompose-"))
    try:
        extract_dir = temp_dir / "pptx_content"
        with zipfile.ZipFile(host_pptx, "r") as zf:
            zf.extractall(extract_dir)

        presentation_path = extract_dir / "ppt" / "presentation.xml"
        presentation_rels_path = extract_dir / "ppt" / "_rels" / "presentation.xml.rels"
        content_types_path = extract_dir / "[Content_Types].xml"
        if not presentation_path.exists() or not presentation_rels_path.exists():
            raise AppendError("invalid PPTX: missing ppt/presentation.xml or its rels")

        presentation_xml = presentation_path.read_text(encoding="utf-8")
        originals = _ordered_slides(presentation_xml, presentation_rels_path)

        keep_indices = [e.index for e in sequence if isinstance(e, KeepSlide)]
        if len(set(keep_indices)) != len(keep_indices):
            raise ValueError("recompose_pptx: duplicate KeepSlide index in sequence")
        for idx in keep_indices:
            if idx < 0 or idx >= len(originals):
                raise ValueError(
                    f"recompose_pptx: KeepSlide({idx}) out of range — deck has "
                    f"{len(originals)} slides"
                )

        layout_rel_target = _pick_blankest_layout(extract_dir)
        slides_dir = extract_dir / "ppt" / "slides"
        next_slide_num = _next_free_slide_number(
            [p.name for p in slides_dir.glob("slide*.xml")] if slides_dir.exists() else []
        )
        next_sld_id = _next_free_sld_id(presentation_xml)

        media_cache: dict[tuple[str, str], str] = {}
        image_exts_used: set[str] = set()

        # 1. Materialise every NewSlide as a package part + presentation rel.
        entries: list[str] = []  # final <p:sldId .../> entries, in order
        for item in sequence:
            if isinstance(item, KeepSlide):
                sld_id, rel_id, _target = originals[item.index]
                entries.append(f'<p:sldId id="{sld_id}" r:id="{rel_id}"/>')
                continue
            slide_num = next_slide_num
            next_slide_num += 1
            write_slide_part_from_svg(
                extract_dir,
                item.svg_path,
                slide_num=slide_num,
                layout_rel_target=layout_rel_target,
                media_cache=media_cache,
                image_exts_used=image_exts_used,
                transition=transition,
                transition_duration=transition_duration,
                verbose=verbose,
            )
            rid = _append_relationship(
                presentation_rels_path, SLIDE_REL_TYPE, f"slides/slide{slide_num}.xml"
            )
            _add_override(
                content_types_path, f"/ppt/slides/slide{slide_num}.xml", SLIDE_CONTENT_TYPE
            )
            entries.append(f'<p:sldId id="{next_sld_id}" r:id="{rid}"/>')
            next_sld_id += 1

        # 2. Rewrite the whole sldIdLst in sequence order.
        presentation_xml = presentation_path.read_text(encoding="utf-8")
        presentation_path.write_text(
            _replace_sld_id_lst(presentation_xml, "".join(entries)), encoding="utf-8"
        )

        # 3. Physically remove originals that were dropped from the deck.
        kept = set(keep_indices)
        for idx, (_sld_id, rel_id, target) in enumerate(originals):
            if idx not in kept:
                remove_slide_parts(extract_dir, rel_id=rel_id, target=target)

        # 4. Content-type defaults for any media the new slides brought in.
        if image_exts_used:
            content_types = content_types_path.read_text(encoding="utf-8")
            for ext in sorted(image_exts_used):
                content_types = _add_default_content_type(
                    content_types, ext, _content_type_for_extension(ext)
                )
            content_types_path.write_text(content_types, encoding="utf-8")

        # 5. Repackage atomically.
        temp_output = temp_dir / "result.pptx"
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in extract_dir.rglob("*"):
                if file_path.is_file():
                    zf.write(file_path, file_path.relative_to(extract_dir))
        shutil.move(str(temp_output), str(output_path))
        return warnings
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------


def _ordered_slides(
    presentation_xml: str, presentation_rels_path: Path
) -> list[tuple[str, str, str]]:
    """Return ``(sld_id, rel_id, target)`` in sldIdLst (deck) order."""
    rel_targets = dict(
        (rel_id, target) for rel_id, target in _existing_slides(presentation_rels_path)
    )
    ordered: list[tuple[str, str, str]] = []
    for match in re.finditer(
        r"<p:sldId\b[^>]*\bid=\"(\d+)\"[^>]*\br:id=\"([^\"]+)\"[^>]*/>", presentation_xml
    ):
        sld_id, rel_id = match.group(1), match.group(2)
        target = rel_targets.get(rel_id)
        if target is not None:
            ordered.append((sld_id, rel_id, target))
    return ordered


def _replace_sld_id_lst(presentation_xml: str, inner: str) -> str:
    """Swap the sldIdLst body for *inner* (handles missing/self-closing lists)."""
    block = f"<p:sldIdLst>{inner}</p:sldIdLst>"
    if re.search(r"<p:sldIdLst\b[^>]*>.*?</p:sldIdLst>", presentation_xml, re.DOTALL):
        return re.sub(
            r"<p:sldIdLst\b[^>]*>.*?</p:sldIdLst>",
            block,
            presentation_xml,
            count=1,
            flags=re.DOTALL,
        )
    if re.search(r"<p:sldIdLst\s*/>", presentation_xml):
        return re.sub(r"<p:sldIdLst\s*/>", block, presentation_xml, count=1)
    if "<p:sldSz" in presentation_xml:
        return presentation_xml.replace("<p:sldSz", f"{block}<p:sldSz", 1)
    raise AppendError("invalid PPTX: presentation.xml has no sldIdLst/sldSz anchor")
