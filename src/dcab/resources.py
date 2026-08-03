"""Paths to bundled generated DCAB fixture data."""

from __future__ import annotations

from pathlib import Path


def bundled_fixture_root() -> Path:
    """Return the installed package's deterministic fixture tree."""

    root = Path(__file__).resolve().parent / "fixtures"
    if not root.is_dir() or root.is_symlink():
        raise RuntimeError("DCAB bundled fixtures are unavailable")
    return root
