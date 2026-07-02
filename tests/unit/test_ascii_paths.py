"""Lint test: every path under src/, tests/, scripts/ is pure ASCII.

This is Track A (filesystem identifiers) enforcement from
ppt-master-analysis/06-bilingual-conventions.md. Failing this test means
someone created a directory or file with non-ASCII characters — rename it.
"""

from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]


def _scan(root: Path) -> list[Path]:
    if not root.exists():
        return []
    bad: list[Path] = []
    for path in root.rglob("*"):
        if any(part.startswith(".") or part == "__pycache__" for part in path.parts):
            continue
        try:
            path.name.encode("ascii")
        except UnicodeEncodeError:
            bad.append(path)
    return bad


def test_src_paths_are_ascii():
    bad = _scan(REPO_ROOT / "src")
    assert not bad, f"Non-ASCII paths under src/: {bad}"


def test_tests_paths_are_ascii():
    bad = _scan(REPO_ROOT / "tests")
    assert not bad, f"Non-ASCII paths under tests/: {bad}"


def test_scripts_paths_are_ascii():
    bad = _scan(REPO_ROOT / "scripts")
    assert not bad, f"Non-ASCII paths under scripts/: {bad}"
