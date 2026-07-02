"""Synchronous, deterministic text edits on a deck asset.

    POST /v1/text-edits
        {pptx_asset_id, edits: [{slide, shape_id, para, new_text, old_text?}]}
        -> {pptx_asset_id: <new revision>, applied, results}

No LLM in the loop — this is the studio canvas's inline text editor, so it
runs in the request (worker thread) and returns the new revision id
immediately. The input asset is preserved (same revision model as
edit-deck jobs).
"""

from __future__ import annotations

import asyncio
import logging
import uuid

from fastapi import APIRouter, HTTPException, status
from pydantic import BaseModel, Field
from sqlalchemy import select

from ...db.models import Asset, AssetKind, Tenant
from ...services.assets import upload_asset
from ...storage import get_default_storage
from ...tools.apply_text_edits import (
    ApplyTextEditsRequest,
    TextEdit,
    apply_text_edits,
)
from ..dependencies import CurrentTenant, DbSession

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/v1/text-edits", tags=["text-edits"])


class TextEditBody(BaseModel):
    slide: int = Field(..., ge=0)
    shape_id: int = Field(..., ge=1)
    para: int = Field(..., ge=0)
    new_text: str
    old_text: str | None = None
    # Table-cell addressing (both set = edit table.cell(row, col)).
    row: int | None = Field(default=None, ge=0)
    col: int | None = Field(default=None, ge=0)


class TextEditsBody(BaseModel):
    pptx_asset_id: uuid.UUID
    edits: list[TextEditBody] = Field(..., min_length=1, max_length=50)
    output_basename: str | None = None


class TextEditsResponse(BaseModel):
    pptx_asset_id: uuid.UUID
    applied: int
    results: list[dict]


@router.post("", response_model=TextEditsResponse, summary="Apply direct text edits")
async def apply_deck_text_edits(
    body: TextEditsBody,
    tenant: CurrentTenant,
    session: DbSession,
) -> TextEditsResponse:
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
            apply_text_edits,
            ApplyTextEditsRequest(
                pptx=content,
                edits=[TextEdit(**e.model_dump()) for e in body.edits],
            ),
        )
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={
                "code": "TEXT_EDIT_FAILED",
                "message": str(exc),
                "message_en": str(exc),
            },
        ) from exc

    results = [r.model_dump() for r in resp.results]
    if resp.applied == 0:
        # Nothing changed — don't forge a new revision; report why per edit.
        return TextEditsResponse(
            pptx_asset_id=body.pptx_asset_id, applied=0, results=results
        )

    tenant_row = (
        await session.execute(select(Tenant).where(Tenant.id == tenant.id))
    ).scalar_one()
    basename = body.output_basename or (asset.original_filename or "deck.pptx").rsplit(
        ".", 1
    )[0]
    upload = await upload_asset(
        session=session,
        storage=storage,
        tenant=tenant_row,
        kind=AssetKind.pptx,
        content=resp.pptx,
        original_filename=f"{basename}.pptx",
        mime_type=(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ),
        project_id=asset.project_id,
    )
    return TextEditsResponse(
        pptx_asset_id=upload.asset.id, applied=resp.applied, results=results
    )
