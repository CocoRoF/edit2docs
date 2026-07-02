"""Application settings for edit2docs server.

Reads from environment variables (and .env) using pydantic-settings.
All env var names use the EDIT2DOCS_ prefix to avoid collisions.

## Standalone vs external-dep modes

The engine runs **standalone by default** — no Postgres, no MinIO, no Redis
required. Set the corresponding env var to opt into external infra:

| Component | Default (standalone) | Opt-in (external) |
|---|---|---|
| Database | `sqlite+aiosqlite:///<data_dir>/edit2docs.db` | `EDIT2DOCS_DATABASE_URL=postgresql+asyncpg://...` |
| Storage  | local filesystem under `<data_dir>/storage/` | `EDIT2DOCS_S3_ENDPOINT_URL=...` + `EDIT2DOCS_S3_BUCKET=...` |
| Queue    | inline asyncio tasks                       | `EDIT2DOCS_REDIS_URL=redis://...` (arq + worker) |

`<data_dir>` defaults to `/data/edit2docs` (writable in the container) and
can be overridden with `EDIT2DOCS_DATA_DIR`.

Auto-bootstrap on startup creates the data dir, runs
`Base.metadata.create_all`, and creates the S3 bucket if missing — no
manual `CREATE DATABASE` / `alembic upgrade` / `mc mb` ever required.
"""

from __future__ import annotations

import secrets
from functools import lru_cache
from pathlib import Path
from typing import Literal

from pydantic import Field, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Top-level server configuration."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        env_prefix="EDIT2DOCS_",
        extra="ignore",
    )

    # Environment
    environment: Literal["development", "staging", "production"] = "development"
    debug: bool = False

    # Server
    host: str = "0.0.0.0"
    port: int = 8000

    # Default language for new projects when none is specified
    default_lang: str = "ko-KR"

    # Standalone data root. SQLite db + local-fs storage live here when no
    # external Postgres / S3 is configured. Must be writable by the engine
    # process. Mount a docker volume here for persistence across restarts.
    data_dir: Path = Path("/data/edit2docs")

    # Database. When unset/empty, defaults to a SQLite file under data_dir.
    # Set to a `postgresql+asyncpg://...` URL to opt into Postgres.
    database_url: str = ""

    # Redis URL for the arq job queue + JobBus pub/sub. When unset/empty the
    # engine runs jobs *inline* via asyncio.create_task — no worker process,
    # no durability across restarts (acceptable for single-instance demos).
    redis_url: str = ""

    # Object storage. When `s3_endpoint_url` AND `s3_bucket` are both set we
    # use the S3 adapter; otherwise files land on the local filesystem under
    # `data_dir/storage/`.
    s3_endpoint_url: str | None = None
    s3_region: str = "us-east-1"
    s3_bucket: str | None = None
    s3_access_key_id: str | None = None
    s3_secret_access_key: str | None = None
    s3_public_base_url: str | None = None

    # HMAC signing key used by the LocalFilesystemStorage adapter to mint
    # presigned-style URLs that hit our own `/v1/raw` endpoint. When unset a
    # random key is generated at startup (NB: a fresh key invalidates any
    # outstanding URLs across restarts; set explicitly for production).
    raw_signing_key: str = ""

    # Auth
    auth_dev_api_key: str | None = Field(
        default=None,
        description="Dev-only static API key. M0 stub; tenant-issued keys later.",
    )

    # Asset TTLs (seconds). Aligned with ppt-master-analysis/04 §4.7.
    asset_ttl_source_seconds: int = 7 * 24 * 60 * 60
    asset_ttl_intermediate_seconds: int = 30 * 24 * 60 * 60
    asset_ttl_pptx_seconds: int = 90 * 24 * 60 * 60

    # Max upload size (bytes)
    max_upload_size_bytes: int = 200 * 1024 * 1024  # 200 MB to match nginx default

    @model_validator(mode="after")
    def _apply_standalone_defaults(self) -> "Settings":
        """Fill in standalone defaults for any unset external-dep settings."""
        # database_url: empty → SQLite under data_dir.
        if not self.database_url:
            self.database_url = f"sqlite+aiosqlite:///{self.data_dir}/edit2docs.db"
        if not self.raw_signing_key:
            self.raw_signing_key = secrets.token_hex(32)
        return self

    # Convenience predicates ---------------------------------------------

    @property
    def uses_postgres(self) -> bool:
        return self.database_url.startswith(
            ("postgresql://", "postgresql+asyncpg://", "postgres://")
        )

    @property
    def uses_sqlite(self) -> bool:
        return self.database_url.startswith("sqlite")

    @property
    def uses_s3_storage(self) -> bool:
        """True when both an S3 endpoint and a bucket name are configured."""
        return bool(self.s3_endpoint_url and self.s3_bucket)

    @property
    def uses_redis_queue(self) -> bool:
        return bool(self.redis_url)


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()


def reset_settings_cache() -> None:
    """Test hook — recompute settings after manipulating env vars."""
    get_settings.cache_clear()
