"""Edit-deck executor: one chat-edit turn on an existing PPTX asset.

Mirrors the generate_deck executor: pull the deck bytes from storage, run
tools.edit_deck with stage events streamed to the job bus, persist the new
revision as a fresh pptx asset (the prior revision stays untouched so the
studio can offer undo by pointing back at it).
"""

from __future__ import annotations

import logging
import uuid
from datetime import datetime, timezone
from typing import Any

from sqlalchemy import select

from ...db.models import Asset, AssetKind, JobEventType, JobKind, JobStatus
from ...services.assets import upload_asset
from ...services.jobs import record_event
from ...storage import get_default_storage
from ...tools import ConvertRequest, StageEvent
from ...tools.edit_deck import ChatTurn, EditDeckRequest, edit_deck
from .generate_deck import _infer_source_type
from .registry import ExecutionContext, register

logger = logging.getLogger(__name__)


@register(JobKind.edit_deck)
async def run_edit_deck(ctx: ExecutionContext) -> None:
    job = ctx.job
    session = ctx.session
    bus = ctx.bus

    job.status = JobStatus.running
    job.started_at = datetime.now(timezone.utc)
    await session.flush()

    params = dict(job.params or {})
    pptx_asset_id = uuid.UUID(params["pptx_asset_id"])
    instruction: str = params["instruction"]
    chat_history: list[dict] = params.get("chat_history", [])
    lang: str = params.get("lang", "ko-KR")
    model: str = params.get("model", "claude-opus-4-7")
    anthropic_api_key: str = params["anthropic_api_key"]

    storage = get_default_storage()

    asset = (
        await session.execute(
            select(Asset).where(
                Asset.id == pptx_asset_id, Asset.tenant_id == job.tenant_id
            )
        )
    ).scalar_one_or_none()
    if asset is None:
        raise RuntimeError(f"deck asset {pptx_asset_id} not found for tenant {job.tenant_id}")
    pptx_bytes = await storage.get_bytes(asset.storage_key)

    # Reference documents attached to this turn.
    convert_reqs: list[ConvertRequest] = []
    for src_id_str in params.get("source_asset_ids", []):
        src_id = uuid.UUID(src_id_str)
        src_asset = (
            await session.execute(
                select(Asset).where(
                    Asset.id == src_id, Asset.tenant_id == job.tenant_id
                )
            )
        ).scalar_one_or_none()
        if src_asset is None:
            raise RuntimeError(
                f"source asset {src_id} not found for tenant {job.tenant_id}"
            )
        convert_reqs.append(
            ConvertRequest(
                source_type=_infer_source_type(src_asset.mime_type),
                content=await storage.get_bytes(src_asset.storage_key),
                original_filename=src_asset.original_filename,
            )
        )

    async def on_event(event: StageEvent) -> None:
        await record_event(
            session=session,
            bus=bus,
            job_id=job.id,
            type=JobEventType.stage if event.stage else JobEventType.progress,
            payload={
                "stage": event.stage,
                "progress": event.progress,
                "message_key": event.message_key,
                "message_vars": event.message_vars,
                "page_index": event.page_index,
            },
        )
        await session.commit()

    resp = await edit_deck(
        EditDeckRequest(
            pptx=pptx_bytes,
            instruction=instruction,
            sources=convert_reqs,
            chat_history=[
                ChatTurn(role=t["role"], content=str(t.get("content", "")))
                for t in chat_history
                if isinstance(t, dict) and t.get("role") in ("user", "assistant")
            ],
            lang=lang,  # type: ignore[arg-type]
            model=model,
            anthropic_api_key=anthropic_api_key,
        ),
        on_event=on_event,
    )

    result: dict[str, Any] = {
        "changed": resp.changed,
        "page_count": resp.page_count,
        "reply": resp.reply,
        "operations": resp.operations,
        "warnings": [{"code": w.code, "message": w.message} for w in resp.warnings],
    }

    if resp.changed:
        from ...db.models import Tenant

        tenant = (
            await session.execute(select(Tenant).where(Tenant.id == job.tenant_id))
        ).scalar_one()
        pptx_upload = await upload_asset(
            session=session,
            storage=storage,
            tenant=tenant,
            kind=AssetKind.pptx,
            content=resp.pptx,
            original_filename=f"{params.get('output_basename', 'deck')}.pptx",
            mime_type=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
            project_id=job.project_id,
        )
        result["pptx_asset_id"] = str(pptx_upload.asset.id)
    else:
        # Question-only turn: the deck is unchanged; keep pointing at the input.
        result["pptx_asset_id"] = str(pptx_asset_id)

    job.result = result
    job.cost = {
        "input_tokens": resp.cost.input_tokens,
        "output_tokens": resp.cost.output_tokens,
        "cache_read_tokens": resp.cost.cache_read_tokens,
        "cache_write_tokens": resp.cost.cache_write_tokens,
        "image_count": resp.cost.image_count,
        "audio_seconds": resp.cost.audio_seconds,
        "duration_seconds": resp.cost.duration_seconds,
    }
    job.status = JobStatus.done
    job.finished_at = datetime.now(timezone.utc)
    await session.commit()
