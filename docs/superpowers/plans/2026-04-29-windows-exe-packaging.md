# Windows .exe Packaging Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `sansa-excel2powerpoint` as a single-file Windows `.exe` built and released by GitHub Actions on tag push, with a Tkinter shell for double-click users and unchanged CLI behavior.

**Architecture:** Add three small modules — `runtime_paths` (frozen-vs-source path resolution), `logging_setup` (per-run log file next to the output `.pptx`), `gui` (Tkinter file picker + success/error dialogs). `__main__.py` branches into the GUI when frozen and no flags were passed; otherwise runs the existing pipeline unchanged. PyInstaller `--onefile --windowed` produces the `.exe`; CI builds, smoke-tests, and attaches a flat zip to the GitHub Release.

**Tech Stack:** Python 3.14, PyInstaller, Tkinter (stdlib), pytest, GitHub Actions on `windows-latest`.

**Spec:** `docs/superpowers/specs/2026-04-29-windows-exe-packaging-design.md`

---

## File Structure

**New files:**
- `requirements.txt` — pip-installable dependency list mirroring `environment.yml` (CI uses pip, not conda).
- `src/runtime_paths.py` — `is_frozen()`, `app_dir()`, `default_template_path()`, `discover_default_excel()`.
- `src/logging_setup.py` — `configure(log_path, *, cli_mode)` configures stdlib `logging` with file + optional stdout handlers.
- `src/gui.py` — Tkinter shell: resolve Excel via drag/auto-discover/picker, run pipeline, show success/error dialogs.
- `tests/__init__.py` — empty.
- `tests/test_runtime_paths.py` — unit tests for path resolution and Excel discovery.
- `tests/test_logging_setup.py` — unit tests for log handler configuration.
- `tests/test_gui_resolve_excel.py` — unit tests for the Excel-resolution decision tree (no Tkinter dialogs).
- `sansa-excel2pptx.spec` — PyInstaller spec checked into repo.
- `.github/workflows/release.yml` — tag-triggered Windows build + Release upload.

**Modified files:**
- `src/builders/template_builder.py:38-40` — replace `DEFAULT_SOURCE` constant with `default_template_path()` call.
- `src/__main__.py` — replace `REPO_ROOT`/`DEFAULT_EXCEL` constants with `runtime_paths` calls, branch into GUI when frozen + no flags, swap `print(...)` for `logging.getLogger(__name__).info(...)`.
- `environment.yml` — bump `python=3.11` → `python=3.14`.
- `README.md` — add "Download for Windows" section, note `requirements.txt`/`environment.yml` are kept in sync manually.
- `.gitignore` — add `dist/`, `build/`, `*.spec.bak`, `tests/__pycache__/` if not already covered.

---

## Task 1: Bootstrap dev tooling (requirements.txt, pytest, env bump)

**Files:**
- Create: `requirements.txt`
- Create: `tests/__init__.py`
- Modify: `environment.yml`
- Modify: `.gitignore`

- [ ] **Step 1: Inspect existing `.gitignore`**

Run: `cat .gitignore`
Note which entries are already present so we don't duplicate them.

- [ ] **Step 2: Create `requirements.txt`**

```
openpyxl>=3.1.5
python-pptx>=1.0.2
lxml
matplotlib>=3.8
```

- [ ] **Step 3: Bump Python in `environment.yml`**

Edit `environment.yml`: change `python=3.11` to `python=3.14`. Add `pytest>=8` and `pyinstaller>=6` under `dependencies` (kept here for parity; CI installs them via pip).

- [ ] **Step 4: Add `tests/__init__.py` (empty file) and update `.gitignore`**

Create empty `tests/__init__.py`. Append to `.gitignore` (only entries not already present):
```
dist/
build/
*.spec.bak
tests/__pycache__/
.pytest_cache/
```

- [ ] **Step 5: Verify pytest discovers the empty test dir**

Run (in your local Python 3.14 env, with pytest installed): `python -m pytest tests/ -v`
Expected: "no tests ran" (exit 5), not an error.

- [ ] **Step 6: Commit**

```bash
git add requirements.txt environment.yml tests/__init__.py .gitignore
git commit -m "chore: add requirements.txt, pytest scaffolding, bump python to 3.14"
```

---

## Task 2: `runtime_paths.py` — path resolution helpers

**Files:**
- Create: `tests/test_runtime_paths.py`
- Create: `src/runtime_paths.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_runtime_paths.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_runtime_paths.py -v`
Expected: 9 errors, all `ModuleNotFoundError: No module named 'src.runtime_paths'`.

- [ ] **Step 3: Implement `src/runtime_paths.py`**

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_runtime_paths.py -v`
Expected: 9 passed.

- [ ] **Step 5: Commit**

```bash
git add src/runtime_paths.py tests/test_runtime_paths.py
git commit -m "feat: add runtime_paths module for frozen-vs-source path resolution"
```

---

## Task 3: Refactor `template_builder.DEFAULT_SOURCE` to use `runtime_paths`

**Files:**
- Modify: `src/builders/template_builder.py:38-40`

- [ ] **Step 1: Replace `DEFAULT_SOURCE` with a function**

Change lines 38-40 of `src/builders/template_builder.py` from:

```python
DEFAULT_SOURCE = (
    Path(__file__).resolve().parents[2] / "labalaba ginuma.pptx"
)
```

to:

```python
from src.runtime_paths import default_template_path


def _default_source() -> Path:
    return default_template_path()


# Backwards-compatible name kept for callers; resolved lazily.
DEFAULT_SOURCE = _default_source()
```

The lazy-resolved `DEFAULT_SOURCE` is fine because it's only read at module-import time and the app_dir is already correct by then in both source and frozen modes.

- [ ] **Step 2: Verify `__main__.py` still imports cleanly**

Run: `python -c "from src.builders.template_builder import DEFAULT_SOURCE; print(DEFAULT_SOURCE)"`
Expected: prints the absolute path to `labalaba ginuma.pptx` in the repo root.

- [ ] **Step 3: Run the existing pipeline end-to-end as a regression check**

Run: `python -m src --excel "excel/labalaba ginuma.xlsx" --output output/regression.pptx`
Expected: `✓ Wrote output/regression.pptx`. Open the file briefly to confirm it looks right.

- [ ] **Step 4: Run all tests**

Run: `python -m pytest tests/ -v`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/builders/template_builder.py
git commit -m "refactor: route template_builder DEFAULT_SOURCE through runtime_paths"
```

---

## Task 4: `logging_setup.py` — per-run logging configuration

**Files:**
- Create: `tests/test_logging_setup.py`
- Create: `src/logging_setup.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_logging_setup.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_logging_setup.py -v`
Expected: 6 errors, `ModuleNotFoundError: No module named 'src.logging_setup'`.

- [ ] **Step 3: Implement `src/logging_setup.py`**

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_logging_setup.py -v`
Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add src/logging_setup.py tests/test_logging_setup.py
git commit -m "feat: add logging_setup module for per-run log file"
```

---

## Task 5: Refactor `__main__.py` to use logging and `runtime_paths`

**Files:**
- Modify: `src/__main__.py`

- [ ] **Step 1: Rewrite `src/__main__.py`**

Replace the entire file contents with:

```python
"""CLI entry point: python -m src --excel ... --date ... --output ..."""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime
from pathlib import Path

from src import logging_setup, runtime_paths
from src.excel_reader import ExcelReader
from src.slide_specs import build_specs
from src.builders.template_builder import TemplateBuilder

log = logging.getLogger(__name__)


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sansa-excel2pptx",
        description="Generate a Sinhala financial PPTX from a monthly Excel workbook.",
    )
    parser.add_argument("--excel", type=Path, default=None,
                        help="Path to the Excel workbook")
    parser.add_argument("--date", type=_parse_date, default=None,
                        help="Target month-end date (YYYY-MM-DD). "
                             "Defaults to the latest date with non-empty data.")
    parser.add_argument("--template", type=Path, default=None,
                        help="Source .pptx whose slide-1 layout (cover image), "
                             "master, and theme are reused for the new deck")
    parser.add_argument("--output", type=Path, default=None,
                        help="Output .pptx path (default: output/<excel>_<date>.pptx)")
    return parser


def _resolve_excel(arg: Path | None) -> Path:
    if arg is not None:
        return arg
    default_excel = runtime_paths.app_dir() / "excel" / "labalaba ginuma.xlsx"
    return default_excel


def _resolve_template(arg: Path | None) -> Path:
    if arg is not None:
        return arg
    return runtime_paths.default_template_path()


def _resolve_output(arg: Path | None, excel: Path, target: date) -> Path:
    if arg is not None:
        return arg
    out_dir = runtime_paths.app_dir() / "output"
    return out_dir / f"{excel.stem}_{target.isoformat()}.pptx"


def run_pipeline(excel: Path, date_arg: date | None, template: Path, output: Path | None,
                 *, cli_mode: bool) -> Path:
    """Run the generation pipeline. Returns the final output path.

    Logging must be configured before any failure can be reported, so we resolve
    paths first, then configure logging, then run.
    """
    reader = ExcelReader(excel)
    target = date_arg or reader.latest_populated_date()
    final_output = _resolve_output(output, excel, target)

    log_path = final_output.with_suffix(final_output.suffix + ".log")
    logging_setup.configure(log_path, cli_mode=cli_mode)

    log.info("Generating deck for %s -> %s", target.isoformat(), final_output)
    specs = build_specs(reader, target)
    log.info("Generating %d spec(s) for %s", len(specs), target.isoformat())
    builder = TemplateBuilder(source_path=template)
    builder.build(specs, reader, target, final_output)
    log.info("Wrote %s", final_output)
    return final_output


def _no_user_flags(argv: list[str]) -> bool:
    """True iff the user invoked the .exe with no flags or with a single .xlsx arg."""
    if len(argv) == 1:
        return True
    if len(argv) == 2 and argv[1].lower().endswith(".xlsx") and Path(argv[1]).exists():
        return True
    return False


def main() -> None:
    if runtime_paths.is_frozen() and _no_user_flags(sys.argv):
        from src import gui
        sys.exit(gui.run())

    args = _build_parser().parse_args()
    excel = _resolve_excel(args.excel)
    template = _resolve_template(args.template)
    run_pipeline(excel, args.date, template, args.output, cli_mode=True)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify CLI still works end-to-end**

Run: `python -m src --excel "excel/labalaba ginuma.xlsx" --output output/refactor-check.pptx`
Expected: stdout shows `Generating deck for ...`, `Wrote ...`. File exists at `output/refactor-check.pptx`. A log file exists at `output/refactor-check.pptx.log`.

- [ ] **Step 3: Verify default-Excel path still works**

Run: `python -m src`
Expected: same behavior as before (uses `excel/labalaba ginuma.xlsx`, writes to `output/labalaba ginuma_<date>.pptx`).

- [ ] **Step 4: Run all tests**

Run: `python -m pytest tests/ -v`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/__main__.py
git commit -m "refactor: route __main__ through runtime_paths and logging_setup"
```

---

## Task 6: `gui.py` Excel-resolution logic + tests

**Files:**
- Create: `tests/test_gui_resolve_excel.py`
- Create: `src/gui.py`

This task only adds the *logic* and unit tests. The actual Tkinter dialogs are wired up in Task 7 (cannot be unit-tested headlessly without extra fixtures, so we verify them manually via the Windows smoke test).

- [ ] **Step 1: Write the failing tests**

Create `tests/test_gui_resolve_excel.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_gui_resolve_excel.py -v`
Expected: 5 errors, `ModuleNotFoundError: No module named 'src.gui'`.

- [ ] **Step 3: Implement `src/gui.py` (logic only — Tkinter wiring in Task 7)**

```python
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
    raise NotImplementedError("wired up in Task 7")
```

`argv` here is the user-provided arg list (i.e., `sys.argv[1:]`), not the full `sys.argv`. The convention keeps the test harness clean.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_gui_resolve_excel.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add src/gui.py tests/test_gui_resolve_excel.py
git commit -m "feat: add gui.resolve_excel decision tree with tests"
```

---

## Task 7: Wire up `gui.run()` with Tkinter dialogs

**Files:**
- Modify: `src/gui.py`

- [ ] **Step 1: Replace `run()` stub with the full implementation**

In `src/gui.py`, replace the body of `run()` (the `raise NotImplementedError(...)` line) with:

```python
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
```

- [ ] **Step 2: Manually verify the GUI logic in source mode**

Source-mode invocation skips `gui.run()` (because `is_frozen()` is False), so end-to-end GUI testing must wait for the frozen build. For now, sanity-check by importing:

Run: `python -c "from src import gui; print(gui.run.__doc__)"`
Expected: prints the docstring without raising.

- [ ] **Step 3: Run all tests**

Run: `python -m pytest tests/ -v`
Expected: all pass.

- [ ] **Step 4: Commit**

```bash
git add src/gui.py
git commit -m "feat: wire up gui.run with Tkinter file picker and dialogs"
```

---

## Task 8: PyInstaller spec file

**Files:**
- Create: `sansa-excel2pptx.spec`

- [ ] **Step 1: Create `sansa-excel2pptx.spec`**

```python
# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for sansa-excel2pptx.

Build with:
    pyinstaller sansa-excel2pptx.spec

Produces a single-file --windowed exe at dist/sansa-excel2pptx.exe.
The template (labalaba ginuma.pptx) is NOT bundled — it ships alongside
the .exe so users can swap in a customised version.
"""

block_cipher = None

a = Analysis(
    ['src/__main__.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'matplotlib.backends.backend_agg',
        'matplotlib.backends.backend_svg',
        'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='sansa-excel2pptx',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # --windowed
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
```

- [ ] **Step 2: Commit**

```bash
git add sansa-excel2pptx.spec
git commit -m "feat: add PyInstaller spec for single-file --windowed exe"
```

Note: building the .exe locally on macOS is not useful (we'd get a Mach-O binary). The first real build happens in CI in Task 9.

---

## Task 9: GitHub Actions workflow

**Files:**
- Create: `.github/workflows/release.yml`

- [ ] **Step 1: Create `.github/workflows/release.yml`**

```yaml
name: release

on:
  push:
    tags:
      - 'v*'

permissions:
  contents: write   # required by softprops/action-gh-release

jobs:
  build-windows:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.14'

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          pip install pyinstaller pytest

      - name: Run unit tests
        run: python -m pytest tests/ -v

      - name: Build exe
        run: pyinstaller sansa-excel2pptx.spec

      - name: Smoke test built exe
        shell: pwsh
        run: |
          & ".\dist\sansa-excel2pptx.exe" `
            --excel "excel\labalaba ginuma.xlsx" `
            --output "smoke.pptx"
          if (-not (Test-Path "smoke.pptx")) {
            Write-Error "smoke.pptx was not produced"
            exit 1
          }
          if ((Get-Item "smoke.pptx").Length -lt 10000) {
            Write-Error "smoke.pptx is suspiciously small"
            exit 1
          }

      - name: Stage release files
        shell: pwsh
        run: |
          New-Item -ItemType Directory -Force -Path staging | Out-Null
          Copy-Item "dist\sansa-excel2pptx.exe" "staging\"
          Copy-Item "labalaba ginuma.pptx" "staging\"
          @"
          sansa-excel2pptx ${{ github.ref_name }}
          ===============================

          Quick start:
          1. Place your monthly Excel workbook (.xlsx) in this folder.
          2. Double-click sansa-excel2pptx.exe.
             - If exactly one .xlsx is here, it will be used automatically.
             - Otherwise a file picker will appear.
             - You can also drag an .xlsx onto sansa-excel2pptx.exe.

          The .pptx is saved next to your Excel input. A log file is
          written alongside the .pptx — attach it when reporting issues.

          CLI (advanced):
            sansa-excel2pptx.exe --excel path\to\workbook.xlsx ^
              --date 2026-03-31 ^
              --output path\to\out.pptx

          Customising the template:
            Replace 'labalaba ginuma.pptx' in this folder with your own
            slide-1 layout. The first slide of that template is reused
            as the cover background.
          "@ | Out-File -Encoding utf8 "staging\README.txt"

      - name: Create zip
        shell: pwsh
        run: |
          Compress-Archive -Path "staging\*" `
            -DestinationPath "sansa-excel2pptx-${{ github.ref_name }}.zip"

      - name: Upload to release
        uses: softprops/action-gh-release@v2
        with:
          files: sansa-excel2pptx-${{ github.ref_name }}.zip
```

- [ ] **Step 2: Commit**

```bash
git add .github/workflows/release.yml
git commit -m "ci: add tag-triggered Windows release workflow"
```

---

## Task 10: README updates

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Add a "Download for Windows" section near the top of `README.md`**

After the introductory paragraph (before "## Setup (one-time)"), insert:

```markdown
## Download for Windows

Pre-built `.exe` releases are published via GitHub Actions on tag pushes
(`v*`). Grab the latest from the [Releases page](../../releases).

The downloaded zip contains:

```
sansa-excel2pptx.exe
labalaba ginuma.pptx          # default template — swap with your own
README.txt
```

Drop your monthly `.xlsx` next to the `.exe` and double-click. See the
included `README.txt` for full usage.
```

- [ ] **Step 2: Add a note about `requirements.txt` near the bottom of `README.md`**

After the "Future work" section (or replace the "PyInstaller Windows packaging" bullet), append:

```markdown
## Dependency files

`environment.yml` (conda) and `requirements.txt` (pip) are kept in sync
manually. Local development uses conda; CI uses pip because PyInstaller
plays poorly with conda environments inside GitHub Actions runners.
```

Remove the obsolete `- PyInstaller Windows packaging` bullet from the "Future work" list.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: document Windows download path and dependency files"
```

---

## Task 11: Tag, observe CI, smoke-test the artifact

This is the manual verification step — there's no code change, just operational verification.

- [ ] **Step 1: Push a pre-release tag**

```bash
git tag v0.1.0-rc1
git push origin v0.1.0-rc1
```

- [ ] **Step 2: Watch the workflow run on GitHub Actions**

Check the run completes green. If it fails:
- "Build exe" failure → likely a hidden import. Read the PyInstaller traceback in the log; add the missing module to `hiddenimports` in the `.spec` file.
- "Smoke test" failure → the exe builds but crashes at runtime. Download the workflow logs and the `smoke.pptx` artifact (if any). Common causes: matplotlib backend, missing `_MEIPASS` data file, `lxml` DLL path.

Fix, commit, and either re-tag (`v0.1.0-rc2`) or delete and retag.

- [ ] **Step 3: Download the published zip from the Release page**

Verify zip layout matches:
```
sansa-excel2pptx.exe
labalaba ginuma.pptx
README.txt
```

- [ ] **Step 4: Test on a Windows machine**

On a Windows box (yours or a VM), test all four flows:
1. Double-click with no `.xlsx` co-located → file picker appears.
2. Double-click with one `.xlsx` co-located → it's used directly.
3. Drag an `.xlsx` onto the `.exe` → it's processed.
4. From `cmd`: `sansa-excel2pptx.exe --excel path\to\book.xlsx --output out.pptx` → CLI path works, stdout visible.

For each, confirm:
- Output `.pptx` is generated.
- Log file `<output>.pptx.log` exists alongside.
- Force a failure (e.g., delete the template, or point at a corrupt xlsx) and verify the error dialog shows the log path.

- [ ] **Step 5: Cut the real release**

Once `rc1` looks good:

```bash
git tag v0.1.0
git push origin v0.1.0
```

The `.exe` is now publicly available on the Releases page.

---

## Spec coverage check (self-review)

| Spec section | Task(s) |
|---|---|
| Distribution & build trigger (tag-triggered, flat zip) | Task 9 |
| Resource discovery (`runtime_paths`) | Task 2, Task 3 |
| Double-click UX (drag / auto-discover / picker) | Task 6, Task 7 |
| Logging next to output, fallback to next-to-exe | Task 4, Task 5, Task 7 |
| GitHub Actions workflow + smoke test | Task 9 |
| `--onefile --windowed` PyInstaller spec | Task 8 |
| Python 3.14 in env + CI | Task 1, Task 9 |
| README updates (download path, dep files) | Task 10 |
| Real-world verification | Task 11 |

All spec sections covered.
