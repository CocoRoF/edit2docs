"""Render-preview tool: PPTX bytes -> self-contained per-slide SVGs.

Deterministic (no LLM). Powers the web studio's slide canvas and the
edit-deck pipeline's "current slide" context: each slide is rendered in
flat inheritance mode (master + layout shapes inlined) with images
base64-embedded, so a browser can display the SVG directly.
"""

from __future__ import annotations

import time

from pydantic import Field

from ..core.pptx_to_svg.converter import ConvertOptions, convert_pptx_to_svg
from ._workspace import temp_workspace
from .types import CostBreakdown, ToolRequest, ToolResponse, WarningEntry

#: OLE/CFB compound-document magic — legacy binary Office (.ppt/.doc/.xls
#: 97-2003). PowerPoint opens these transparently, so a file that "works"
#: there can still be a legacy .ppt merely renamed to .pptx — which is NOT
#: a zip and cannot be a real OOXML .pptx.
_OLE_MAGIC = b"\xd0\xcf\x11\xe0"
_ZIP_MAGIC = b"PK\x03\x04"


def _diagnose_non_pptx(data: bytes) -> str | None:
    """A precise, actionable reason when *data* isn't a valid .pptx (a zip),
    or None to fall back to the generic message. Bilingual (ko + en)."""
    if not data:
        return "빈 파일이 업로드되었습니다. (the uploaded file is empty)"
    head = data[:8]
    if head[:4] == _OLE_MAGIC:
        return (
            "이 파일은 구형 PowerPoint 형식(.ppt, 97–2003 이진 형식)이라 미리보기·편집이 "
            "불가능합니다. PowerPoint에서 [파일 → 다른 이름으로 저장 → 파일 형식: "
            "PowerPoint 프레젠테이션 (*.pptx)]로 다시 저장한 뒤 업로드하세요. "
            "(this is a legacy binary .ppt, not an OOXML .pptx — re-save as .pptx)"
        )
    if head[:1] == b"<" or head[:5].lower() == b"<html":
        return (
            "업로드된 것이 문서가 아니라 HTML/오류 페이지입니다. 올바른 .pptx 파일을 "
            "업로드했는지 확인하세요. (received HTML, not a document)"
        )
    if head[:4] != _ZIP_MAGIC:
        return (
            "이 파일은 올바른 .pptx(OOXML zip) 형식이 아닙니다 — 파일이 손상되었거나 "
            "확장자만 .pptx로 바뀐 다른 형식일 수 있습니다. PowerPoint에서 .pptx로 다시 "
            "저장해 보세요. (not a valid .pptx / OOXML zip)"
        )
    return None


class RenderPreviewRequest(ToolRequest):
    pptx: bytes = Field(..., description="The PPTX package to render.")
    max_slides: int = Field(
        default=100,
        ge=1,
        le=500,
        description="Safety cap; decks longer than this are truncated with a warning.",
    )


class SlidePreview(ToolResponse):
    index: int = Field(..., description="0-based slide position.")
    svg: str = Field(..., description="Self-contained SVG (images embedded).")


class RenderPreviewResponse(ToolResponse):
    slides: list[SlidePreview]
    width_px: float
    height_px: float
    page_count: int = Field(..., description="Total slides in the deck (pre-truncation).")
    cost: CostBreakdown
    warnings: list[WarningEntry] = Field(default_factory=list)


def render_preview(req: RenderPreviewRequest) -> RenderPreviewResponse:
    """Convert every slide to a flat, self-contained SVG.

    Raises:
        ValueError: the bytes are not a readable PPTX (bilingual message).
    """
    started = time.perf_counter()
    warnings: list[WarningEntry] = []

    with temp_workspace(prefix="edit2docs-preview-") as ws:
        pptx_path = ws / "deck.pptx"
        pptx_path.write_bytes(req.pptx)
        out_dir = ws / "svg"
        out_dir.mkdir()
        try:
            result = convert_pptx_to_svg(
                pptx_path,
                out_dir,
                ConvertOptions(embed_images=True, inheritance_mode="flat"),
            )
        except Exception as exc:
            hint = _diagnose_non_pptx(req.pptx)
            if hint is not None:
                raise ValueError(hint) from exc
            raise ValueError(
                f"PPTX could not be rendered for preview: {exc}. "
                "PPTX 파일을 미리보기로 변환할 수 없습니다 — 올바른 .pptx 파일인지 확인하세요."
            ) from exc

    # Pure flat mode populates `slides` with the flat view.
    artifacts = result.slides
    page_count = len(artifacts)
    if page_count > req.max_slides:
        warnings.append(
            WarningEntry(
                code="preview_truncated",
                message=(
                    f"Deck has {page_count} slides; preview truncated to the "
                    f"first {req.max_slides}."
                ),
                detail={"page_count": page_count, "max_slides": req.max_slides},
            )
        )
        artifacts = artifacts[: req.max_slides]

    return RenderPreviewResponse(
        slides=[
            SlidePreview(index=i, svg=a.svg) for i, a in enumerate(artifacts)
        ],
        width_px=result.canvas_px[0],
        height_px=result.canvas_px[1],
        page_count=page_count,
        cost=CostBreakdown(duration_seconds=time.perf_counter() - started),
        warnings=warnings,
    )
