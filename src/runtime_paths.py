"""Path resolution that's aware of PyInstaller --onefile bundling.

In a PyInstaller --onefile build, ``__file__`` and ``sys._MEIPASS`` point to
a temp extraction directory that disappears between runs. The actual install
directory (where the user dropped the template and any .xlsx) is the parent
of ``sys.executable``.
"""
from __future__ import annotations

import sys
from pathlib import Path

TEMPLATE_FILENAME = "labalaba ginuma.pptx"


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def app_dir() -> Path:
    """Directory containing the .exe (frozen) or repo root (source)."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def default_template_path() -> Path:
    return app_dir() / TEMPLATE_FILENAME


def discover_default_excel() -> Path | None:
    """In frozen mode, return the sole co-located .xlsx if exactly one exists."""
    if not is_frozen():
        return None
    candidates = sorted(app_dir().glob("*.xlsx"))
    return candidates[0] if len(candidates) == 1 else None
