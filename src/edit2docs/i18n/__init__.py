"""Bilingual message catalog (ko / en) for edit2docs user-facing strings.

See ppt-master-analysis/06-bilingual-conventions.md for the policy:
- Track A (filesystem, code, DB, API path): English ASCII only
- Track B (UI, errors, MCP descriptions): bilingual ko + en via this catalog
- Track C (user content): preserved as-is
"""

from .catalog import (
    DEFAULT_LOCALE,
    FALLBACK_LOCALE,
    MessageCatalog,
    default_catalog,
    normalize_locale,
    t,
)

__all__ = [
    "DEFAULT_LOCALE",
    "FALLBACK_LOCALE",
    "MessageCatalog",
    "default_catalog",
    "normalize_locale",
    "t",
]
