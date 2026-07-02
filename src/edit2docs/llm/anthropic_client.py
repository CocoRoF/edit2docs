"""Anthropic SDK wrapper with BYOK + prompt caching + retry.

Used by the Strategist and Executor tools. BYOK = the caller passes their own
Anthropic API key per request; we never hold a server-wide key.

Prompt caching: the long system prompts (strategist.ko.md is ~500 lines,
executor-base.ko.md is ~700 lines) are sent with `cache_control: ephemeral`
so subsequent calls within the 5-minute TTL only pay for the user delta.
This is critical for cost — per-page Executor calls reuse the same system
prompt N times and would be ~$5+/deck without caching.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any

import httpx

logger = logging.getLogger(__name__)

# Default model — Opus 4.7 (1M context) per ppt-master-analysis recommendation.
# Callers may override per-request when speed matters more than ceiling.
DEFAULT_MODEL = "claude-opus-4-7"

# Retries on transient failures (429, 502, 503, 504). Exponential backoff with
# jitter is handled by the Anthropic SDK; we only configure the retry count.
DEFAULT_MAX_RETRIES = 3


@dataclass
class LLMUsage:
    """Per-call token accounting, mirrored from anthropic.types.Usage."""

    input_tokens: int = 0
    output_tokens: int = 0
    cache_read_tokens: int = 0
    cache_write_tokens: int = 0

    @property
    def billable_tokens(self) -> int:
        """Tokens that cost real money (cache reads are billed at 10%)."""
        return self.input_tokens + self.output_tokens

    def __add__(self, other: "LLMUsage") -> "LLMUsage":
        return LLMUsage(
            input_tokens=self.input_tokens + other.input_tokens,
            output_tokens=self.output_tokens + other.output_tokens,
            cache_read_tokens=self.cache_read_tokens + other.cache_read_tokens,
            cache_write_tokens=self.cache_write_tokens + other.cache_write_tokens,
        )


@dataclass
class LLMResult:
    """One non-streaming completion."""

    text: str
    usage: LLMUsage
    model: str
    stop_reason: str | None


class AnthropicClient:
    """Thin BYOK-friendly wrapper around `anthropic.AsyncAnthropic`.

    The SDK is imported lazily so that test code (which never calls `complete`)
    doesn't need the `anthropic` package installed.
    """

    def __init__(
        self,
        api_key: str,
        *,
        model: str = DEFAULT_MODEL,
        max_retries: int = DEFAULT_MAX_RETRIES,
        timeout_seconds: float = 600.0,
        base_url: str | None = None,
    ):
        if not api_key:
            raise ValueError("api_key is required (BYOK — caller must pass their own key)")
        self._api_key = api_key
        self._model = model
        self._max_retries = max_retries
        self._timeout_seconds = timeout_seconds
        self._base_url = base_url
        self._client = None  # lazy

    def _ensure_client(self):
        if self._client is None:
            try:
                from anthropic import AsyncAnthropic  # type: ignore[import-not-found]
            except ImportError as exc:
                raise RuntimeError(
                    "`anthropic` package not installed. Install with: "
                    "uv pip install anthropic (or pip install anthropic)"
                ) from exc
            kwargs: dict[str, Any] = {
                "api_key": self._api_key,
                "max_retries": self._max_retries,
                "timeout": httpx.Timeout(self._timeout_seconds),
            }
            if self._base_url:
                kwargs["base_url"] = self._base_url
            self._client = AsyncAnthropic(**kwargs)
        return self._client

    async def complete(
        self,
        system_prompt: str,
        user_message: str,
        *,
        max_output_tokens: int = 8192,
        temperature: float | None = None,
        cache_system: bool = True,
        model: str | None = None,
    ) -> LLMResult:
        """One-shot completion with optional system-prompt caching.

        `temperature` is optional and intentionally omitted from the SDK
        call when None — recent Claude models (Opus 4.5+) reject the
        parameter outright with `temperature is deprecated for this model`.
        """
        client = self._ensure_client()
        system_blocks: list[dict[str, Any]] = [{"type": "text", "text": system_prompt}]
        if cache_system:
            system_blocks[0]["cache_control"] = {"type": "ephemeral"}

        create_kwargs: dict[str, Any] = {
            "model": model or self._model,
            "max_tokens": max_output_tokens,
            "system": system_blocks,
            "messages": [{"role": "user", "content": user_message}],
        }
        if temperature is not None:
            create_kwargs["temperature"] = temperature

        response = await client.messages.create(**create_kwargs)
        usage = _extract_usage(response)
        text = _extract_text(response)
        return LLMResult(
            text=text,
            usage=usage,
            model=model or self._model,
            stop_reason=getattr(response, "stop_reason", None),
        )


def _extract_text(response: Any) -> str:
    """Concatenate text blocks from a Messages response."""
    parts: list[str] = []
    for block in getattr(response, "content", []) or []:
        block_type = getattr(block, "type", None)
        if block_type == "text":
            parts.append(getattr(block, "text", ""))
    return "".join(parts)


def _extract_usage(response: Any) -> LLMUsage:
    usage = getattr(response, "usage", None)
    if usage is None:
        return LLMUsage()
    return LLMUsage(
        input_tokens=getattr(usage, "input_tokens", 0) or 0,
        output_tokens=getattr(usage, "output_tokens", 0) or 0,
        cache_read_tokens=getattr(usage, "cache_read_input_tokens", 0) or 0,
        cache_write_tokens=getattr(usage, "cache_creation_input_tokens", 0) or 0,
    )
