"""``GET /v1/models`` — the model picker's source of truth.

The hosted studio's model dropdown was hardcoded and drifted out of date. This
route serves it **live**: given the caller's BYOK Anthropic key (via the
``X-Anthropic-API-Key`` header, same channel as generate/edit), it proxies the
Anthropic Models API and returns the models that key can actually use, newest
first. With no key — or if the upstream call fails — it returns a curated
fallback list so the UI always has something sensible to show.

Keys are never stored; the header value is used for the single upstream call
and discarded.
"""

from __future__ import annotations

import logging
from typing import Annotated, Any

from fastapi import APIRouter, Header

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/v1", tags=["models"])

# Curated fallback — current text models, offered when no working key is
# available. Kept intentionally small; a live fetch supersedes it entirely.
# The default (index 0) matches the engine's DEFAULT_MODEL so offline
# behavior is unchanged from before this route existed.
_FALLBACK: list[dict[str, str]] = [
    {"id": "claude-opus-4-7", "display_name": "Claude Opus 4.7"},
    {"id": "claude-opus-4-8", "display_name": "Claude Opus 4.8"},
    {"id": "claude-sonnet-4-6", "display_name": "Claude Sonnet 4.6"},
    {"id": "claude-sonnet-5", "display_name": "Claude Sonnet 5"},
    {"id": "claude-haiku-4-5", "display_name": "Claude Haiku 4.5"},
]

# When a live list comes back, prefer one of these (first match) as the
# default selection; else fall back to the newest returned model.
_DEFAULT_PREFERENCE: tuple[str, ...] = (
    "claude-opus-4-8",
    "claude-opus-4-7",
    "claude-opus-5",
    "claude-sonnet-5",
    "claude-sonnet-4-6",
)

# Only advertise chat/completion models. Anthropic model ids are all
# ``claude-*``; this guards against any non-chat entry sneaking in.
_ID_PREFIX = "claude-"


async def _fetch_live(key: str) -> list[dict[str, str]]:
    """The models this key can use, newest first. Raises on any failure."""
    from anthropic import AsyncAnthropic

    client = AsyncAnthropic(api_key=key, max_retries=0, timeout=12.0)
    try:
        rows: list[tuple[Any, dict[str, str]]] = []
        async for m in client.models.list():
            mid = getattr(m, "id", "") or ""
            if not mid.startswith(_ID_PREFIX):
                continue
            rows.append(
                (
                    getattr(m, "created_at", None),
                    {"id": mid, "display_name": getattr(m, "display_name", None) or mid},
                )
            )
    finally:
        await client.close()
    # Newest first when created_at is available; stable otherwise.
    rows.sort(key=lambda r: str(r[0] or ""), reverse=True)
    return [row for _, row in rows]


def _pick_default(models: list[dict[str, str]]) -> str:
    ids = {m["id"] for m in models}
    for preferred in _DEFAULT_PREFERENCE:
        if preferred in ids:
            return preferred
    return models[0]["id"] if models else _FALLBACK[0]["id"]


@router.get("/models")
async def list_models(
    x_anthropic_api_key: Annotated[
        str | None, Header(alias="X-Anthropic-API-Key")
    ] = None,
) -> dict[str, Any]:
    """List selectable models. ``source`` is ``live`` or ``fallback``."""
    key = (x_anthropic_api_key or "").strip()
    if key:
        try:
            models = await _fetch_live(key)
            if models:
                return {
                    "models": models,
                    "default": _pick_default(models),
                    "source": "live",
                }
        except Exception as exc:  # network / auth / SDK — degrade gracefully
            logger.info("live model fetch failed (%s); serving fallback", type(exc).__name__)
    return {
        "models": list(_FALLBACK),
        "default": _FALLBACK[0]["id"],
        "source": "fallback",
    }
