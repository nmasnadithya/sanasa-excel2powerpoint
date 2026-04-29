"""Configures stdlib logging for the app.

File handler: always attached, writes to ``log_path`` at DEBUG level.
Stream handler: only attached in CLI mode (writes INFO to stdout).
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path

_FILE_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"
_STREAM_FORMAT = "%(message)s"


def configure(log_path: Path, *, cli_mode: bool) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(_FILE_FORMAT))
    root.addHandler(file_handler)

    if cli_mode:
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setLevel(logging.INFO)
        stream_handler.setFormatter(logging.Formatter(_STREAM_FORMAT))
        root.addHandler(stream_handler)
