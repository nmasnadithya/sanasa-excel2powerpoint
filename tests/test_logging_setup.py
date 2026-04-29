"""Tests for src.logging_setup."""
from __future__ import annotations

import logging
from pathlib import Path

import pytest

from src import logging_setup


@pytest.fixture(autouse=True)
def _reset_root_logger():
    """Each test gets a clean root logger."""
    root = logging.getLogger()
    saved_handlers = root.handlers[:]
    saved_level = root.level
    root.handlers = []
    yield
    root.handlers = saved_handlers
    root.level = saved_level


def test_configure_attaches_file_handler(tmp_path: Path):
    log_path = tmp_path / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=False)

    file_handlers = [
        h for h in logging.getLogger().handlers
        if isinstance(h, logging.FileHandler)
    ]
    assert len(file_handlers) == 1
    assert Path(file_handlers[0].baseFilename) == log_path


def test_configure_cli_mode_adds_stream_handler(tmp_path: Path):
    log_path = tmp_path / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=True)

    stream_handlers = [
        h for h in logging.getLogger().handlers
        if isinstance(h, logging.StreamHandler)
        and not isinstance(h, logging.FileHandler)
    ]
    assert len(stream_handlers) == 1


def test_configure_non_cli_mode_no_stream_handler(tmp_path: Path):
    log_path = tmp_path / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=False)

    stream_handlers = [
        h for h in logging.getLogger().handlers
        if isinstance(h, logging.StreamHandler)
        and not isinstance(h, logging.FileHandler)
    ]
    assert stream_handlers == []


def test_log_messages_land_in_file(tmp_path: Path):
    log_path = tmp_path / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=False)

    logging.getLogger("test").info("hello world")
    for handler in logging.getLogger().handlers:
        handler.flush()

    contents = log_path.read_text(encoding="utf-8")
    assert "hello world" in contents


def test_configure_creates_parent_dir(tmp_path: Path):
    log_path = tmp_path / "nested" / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=False)
    assert log_path.parent.exists()


def test_exception_traceback_recorded(tmp_path: Path):
    log_path = tmp_path / "out.pptx.log"
    logging_setup.configure(log_path, cli_mode=False)

    log = logging.getLogger("test")
    try:
        raise ValueError("boom")
    except ValueError:
        log.exception("caught")
    for handler in logging.getLogger().handlers:
        handler.flush()

    contents = log_path.read_text(encoding="utf-8")
    assert "boom" in contents
    assert "Traceback" in contents
