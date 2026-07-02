"""Fail if any file or directory under tracked roots has non-ASCII characters.

Enforces Track A of the bilingual conventions (see
ppt-master-analysis/06-bilingual-conventions.md): all filesystem identifiers
must be pure ASCII. Run in CI to prevent regressions.

Usage:
    python3 scripts/lint_ascii_paths.py            # scan default roots
    python3 scripts/lint_ascii_paths.py src tests  # explicit roots
"""

from __future__ import annotations

import sys
from pathlib import Path

DEFAULT_ROOTS = ["src", "tests", "scripts"]


def has_non_ascii(name: str) -> bool:
    try:
        name.encode("ascii")
        return False
    except UnicodeEncodeError:
        return True


def scan(root: Path) -> list[Path]:
    offenders: list[Path] = []
    if not root.exists():
        return offenders
    for path in root.rglob("*"):
        # Skip dotted dirs (e.g. .git, __pycache__) — they belong to tools.
        if any(part.startswith(".") or part == "__pycache__" for part in path.parts):
            continue
        if has_non_ascii(path.name):
            offenders.append(path)
    return offenders


def main(argv: list[str]) -> int:
    roots = [Path(p) for p in (argv[1:] or DEFAULT_ROOTS)]
    all_offenders: list[Path] = []
    for root in roots:
        all_offenders.extend(scan(root))
    if all_offenders:
        sys.stderr.write("Non-ASCII paths found (violates Track A convention):\n")
        for path in all_offenders:
            sys.stderr.write(f"  {path}\n")
        sys.stderr.write(
            "\nRename to ASCII identifiers and update references. "
            "See ppt-master-analysis/06-bilingual-conventions.md.\n"
        )
        return 1
    print(f"OK: all paths under {', '.join(str(r) for r in roots)} are ASCII.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
