"""Tests for src.gui Excel resolution decision tree."""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from src import gui


def test_argv_xlsx_is_used_when_exists(tmp_path: Path):
    workbook = tmp_path / "monthly.xlsx"
    workbook.touch()
    result = gui.resolve_excel(argv=[str(workbook)], picker=lambda: None)
    assert result == workbook


def test_argv_xlsx_ignored_when_does_not_exist(tmp_path: Path):
    """argv[1] must point to an existing file; otherwise fall through."""
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    nonexistent = tmp_path / "ghost.xlsx"
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        result = gui.resolve_excel(argv=[str(nonexistent)], picker=lambda: None)
    assert result is None


def test_auto_discover_when_no_argv(tmp_path: Path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    workbook = tmp_path / "monthly.xlsx"
    workbook.touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        result = gui.resolve_excel(argv=[], picker=lambda: None)
    assert result == workbook


def test_picker_called_when_no_argv_and_multiple_xlsx(tmp_path: Path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    (tmp_path / "a.xlsx").touch()
    (tmp_path / "b.xlsx").touch()
    picker_chose = tmp_path / "a.xlsx"
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        result = gui.resolve_excel(argv=[], picker=lambda: picker_chose)
    assert result == picker_chose


def test_picker_returning_none_yields_none(tmp_path: Path):
    """User cancelled the picker."""
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        result = gui.resolve_excel(argv=[], picker=lambda: None)
    assert result is None
