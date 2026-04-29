"""Tests for src.runtime_paths."""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

from src import runtime_paths


def test_is_frozen_false_in_source_mode():
    assert runtime_paths.is_frozen() is False


def test_is_frozen_true_when_sys_frozen_set():
    with patch.object(sys, "frozen", True, create=True):
        assert runtime_paths.is_frozen() is True


def test_app_dir_in_source_mode_is_repo_root():
    repo_root = Path(__file__).resolve().parents[1]
    assert runtime_paths.app_dir() == repo_root


def test_app_dir_in_frozen_mode_is_executable_parent(tmp_path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        assert runtime_paths.app_dir() == tmp_path


def test_default_template_path_uses_app_dir(tmp_path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        assert runtime_paths.default_template_path() == tmp_path / "labalaba ginuma.pptx"


def test_discover_default_excel_returns_none_in_source_mode():
    assert runtime_paths.discover_default_excel() is None


def test_discover_default_excel_returns_none_when_no_xlsx(tmp_path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        assert runtime_paths.discover_default_excel() is None


def test_discover_default_excel_returns_sole_xlsx(tmp_path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    workbook = tmp_path / "monthly.xlsx"
    workbook.touch()
    (tmp_path / "labalaba ginuma.pptx").touch()  # template, ignored
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        assert runtime_paths.discover_default_excel() == workbook


def test_discover_default_excel_returns_none_when_multiple_xlsx(tmp_path):
    fake_exe = tmp_path / "sansa-excel2pptx.exe"
    fake_exe.touch()
    (tmp_path / "a.xlsx").touch()
    (tmp_path / "b.xlsx").touch()
    with patch.object(sys, "frozen", True, create=True), \
         patch.object(sys, "executable", str(fake_exe)):
        assert runtime_paths.discover_default_excel() is None
