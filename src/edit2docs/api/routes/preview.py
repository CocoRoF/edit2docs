"""Synchronous PPTX preview: every slide rendered to a self-contained SVG.

    POST /v1/preview   {pptx_asset_id} -> {slides: [{index, svg}], ...}

Deterministic and LLM-free, so it runs inline in the request (in a worker
thread) rather than through the job queue — the studio needs previews
immediately after upload and after every edit turn.
"""

from __future__ import annotations

import asyncio
import logging
import uuid

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel
from sqlalchemy import select

from ...db.models import Asset
from ...storage import get_default_storage
from ...tools import RenderPreviewRequest, render_preview
from ..dependencies import CurrentTenant, DbSession

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/v1/preview", tags=["preview"])


class PreviewBody(BaseModel):
    pptx_asset_id: uuid.UUID


class PreviewSlide(BaseModel):
    index: int
    svg: str


class PreviewResponse(BaseModel):
    slides: list[PreviewSlide]
    width_px: float
    height_px: float
    page_count: int
    warnings: list[dict]


@router.post("", response_model=PreviewResponse, summary="Render deck slides to SVG")
async def preview_deck(
    body: PreviewBody,
    tenant: CurrentTenant,
    session: DbSession,
) -> PreviewResponse:
    asset = (
        await session.execute(
            select(Asset).where(
                Asset.id == body.pptx_asset_id, Asset.tenant_id == tenant.id
            )
        )
    ).scalar_one_or_none()
    if asset is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={
                "code": "ASSET_NOT_FOUND",
                "message": f"에셋 {body.pptx_asset_id} 를 찾을 수 없습니다.",
                "message_en": f"Asset {body.pptx_asset_id} not found.",
            },
        )

    storage = get_default_storage()
    content = await storage.get_bytes(asset.storage_key)
    try:
        resp = await asyncio.to_thread(
            render_preview, RenderPreviewRequest(pptx=content)
        )
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={
                "code": "PREVIEW_RENDER_FAILED",
                "message": str(exc),
                "message_en": str(exc),
            },
        ) from exc

    return PreviewResponse(
        slides=[PreviewSlide(index=s.index, svg=s.svg) for s in resp.slides],
        width_px=resp.width_px,
        height_px=resp.height_px,
        page_count=resp.page_count,
        warnings=[{"code": w.code, "message": w.message} for w in resp.warnings],
    )
