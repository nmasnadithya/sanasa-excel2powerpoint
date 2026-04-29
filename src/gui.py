"""Tkinter shell for double-click users.

Frozen-mode entry point. Resolves the Excel input via three rules in order:
1. argv[1] if it points to an existing .xlsx (drag-onto-icon).
2. Sole .xlsx co-located with the .exe (auto-discover).
3. File-open dialog.

Then runs the pipeline and shows a success or error message box.
"""
from __future__ import annotations

import logging
import os
import sys
from pathlib import Path
from typing import Callable, Optional

from src import runtime_paths

log = logging.getLogger(__name__)


def resolve_excel(argv: list[str], picker: Callable[[], Optional[Path]]) -> Optional[Path]:
    """Run the resolution decision tree. Returns the path or None if cancelled."""
    if argv:
        candidate = Path(argv[0])
        if candidate.suffix.lower() == ".xlsx" and candidate.exists():
            return candidate

    discovered = runtime_paths.discover_default_excel()
    if discovered is not None:
        return discovered

    return picker()


def run() -> int:
    """Frozen-mode entry point. Returns process exit code."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    # Hide the empty root window — we only want dialogs.
    root = tk.Tk()
    root.withdraw()

    def _picker() -> Optional[Path]:
        chosen = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=str(runtime_paths.app_dir()),
        )
        return Path(chosen) if chosen else None

    excel = resolve_excel(argv=sys.argv[1:], picker=_picker)
    if excel is None:
        # User cancelled — exit silently.
        return 0

    template = runtime_paths.default_template_path()
    if not template.exists():
        messagebox.showerror(
            "Missing template",
            f"Could not find {template.name} next to the application.\n\n"
            f"Expected location:\n{template}",
        )
        return 1

    # Defer pipeline import until after dialogs so a missing-template error
    # doesn't trigger a slow openpyxl/matplotlib import.
    from src.__main__ import run_pipeline

    try:
        output = run_pipeline(
            excel=excel,
            date_arg=None,
            template=template,
            output=None,
            cli_mode=False,
        )
    except Exception as exc:
        log.exception("Pipeline failed")
        # Best-effort log path — may not exist if failure was before logging_setup ran.
        log_path = _best_effort_log_path(excel)
        _show_error_dialog(messagebox, exc, log_path)
        return 1

    log_path = output.with_suffix(output.suffix + ".log")
    if messagebox.askyesno(
        "Done",
        f"Saved to:\n{output}\n\nOpen the output folder?",
    ):
        os.startfile(output.parent)
    return 0


def _best_effort_log_path(excel: Path) -> Optional[Path]:
    """When the pipeline fails before logging_setup ran, write a fallback log."""
    fallback = runtime_paths.app_dir() / "sansa-excel2pptx-error.log"
    # If logging is already configured, find the first FileHandler's path.
    for handler in logging.getLogger().handlers:
        if isinstance(handler, logging.FileHandler):
            return Path(handler.baseFilename)
    # Ensure fallback exists so we can point the user at it.
    try:
        with fallback.open("a", encoding="utf-8") as f:
            f.write(f"Pipeline failed for excel={excel}\n")
    except OSError:
        return None
    return fallback


def _show_error_dialog(messagebox, exc: BaseException, log_path: Optional[Path]) -> None:
    title = "Generation failed"
    body = f"{type(exc).__name__}: {exc}"
    if log_path is not None:
        body += f"\n\nLog file:\n{log_path}"
        if messagebox.askyesno(title, body + "\n\nOpen log folder?"):
            os.startfile(log_path.parent)
    else:
        messagebox.showerror(title, body)
