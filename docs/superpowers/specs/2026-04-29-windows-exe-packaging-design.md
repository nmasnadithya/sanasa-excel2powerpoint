# Windows Executable Packaging — Design

**Date:** 2026-04-29
**Status:** Approved (pending user spec review)

## Goal

Ship `sansa-excel2powerpoint` as a standalone Windows `.exe` so a non-technical user can generate the PowerPoint deck without installing Python or conda. Build is automated via GitHub Actions on tag push.

## End-user experience

User downloads `sansa-excel2pptx-vX.Y.Z.zip` from the repo's GitHub Releases page and extracts it. The extracted folder contains:

```
sansa-excel2pptx.exe
labalaba ginuma.pptx          # default template, swappable
README.txt                    # short usage note
```

Three ways to use it:

1. **Drag-and-drop** — drag an `.xlsx` onto `sansa-excel2pptx.exe`. The deck is generated next to the input Excel.
2. **Double-click with co-located workbook** — if exactly one `.xlsx` sits in the same folder as the `.exe` (excluding `labalaba ginuma.pptx`), double-clicking the `.exe` uses that file directly.
3. **Double-click without a workbook** — a Windows file-open dialog appears asking for an `.xlsx`. Cancelling exits silently.

After processing, a "Done — saved to ..." message box appears. On error, an error dialog shows the message plus the log file path with an "Open log folder" button.

CLI usage from `cmd`/PowerShell continues to work exactly as today — same `--excel`, `--date`, `--template`, `--output` flags, same stdout output. The GUI flow is bypassed when any flag is passed.

## Distribution model

- Trigger: `push` of a tag matching `v*` (e.g. `v1.0.0`).
- Workflow builds on `windows-latest`, produces the zip, attaches it to the GitHub Release for the tag.
- No CI on `main` pushes. Local PyInstaller runs are fine for pre-tag testing.

## Architecture

### New files

| File | Purpose |
|---|---|
| `src/runtime_paths.py` | Resolves `app_dir`, default template path, default Excel discovery. Isolates frozen-vs-source logic. |
| `src/gui.py` | Tkinter shell used only in frozen mode when no CLI flags are present. File picker, success/error dialogs, log-folder opener. |
| `src/logging_setup.py` | Configures Python `logging` with a per-run file handler (next to output) and an optional stdout handler (CLI mode). |
| `sansa-excel2pptx.spec` | PyInstaller build spec (checked into repo). |
| `requirements.txt` | pip-installable dependency list mirroring `environment.yml`. Used by CI. |
| `.github/workflows/release.yml` | Tag-triggered Windows build + Release upload. |

### Modified files

| File | Change |
|---|---|
| `src/__main__.py` | Replace `REPO_ROOT` / `DEFAULT_EXCEL` constants with `runtime_paths` calls. Branch into `gui.run()` when frozen and no CLI flags passed. Replace `print(...)` with `logging.getLogger(__name__).info(...)`. |
| `src/builders/template_builder.py` | Replace `DEFAULT_SOURCE` constant with a function call into `runtime_paths.default_template_path()`. |
| `environment.yml` | Bump `python=3.11` → `python=3.14` to match CI and local environment. |
| `README.md` | Add a "Download a Windows build" section pointing to Releases. Note that `requirements.txt` and `environment.yml` are kept in sync manually. |

## Component details

### `src/runtime_paths.py`

```python
def is_frozen() -> bool:
    return getattr(sys, "frozen", False)

def app_dir() -> Path:
    """Directory containing the .exe (frozen) or repo root (source)."""
    if is_frozen():
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]

def default_template_path() -> Path:
    return app_dir() / "labalaba ginuma.pptx"

def discover_default_excel() -> Path | None:
    """In frozen mode, return the sole co-located .xlsx if exactly one exists."""
    if not is_frozen():
        return None
    candidates = [p for p in app_dir().glob("*.xlsx")]
    return candidates[0] if len(candidates) == 1 else None
```

**Why anchor on `Path(sys.executable).parent` and not `__file__`:** in PyInstaller `--onefile` mode, `__file__` and `sys._MEIPASS` point to a temp extraction directory that disappears between runs. The actual install dir (where the user dropped `labalaba ginuma.pptx` and any `.xlsx`) is `Path(sys.executable).parent`.

### `src/gui.py`

Pure Tkinter (stdlib, no new dependency). Public surface:

```python
def run() -> int:
    """Frozen-mode entry point. Returns process exit code."""
```

Flow:
1. Resolve Excel: `argv[1]` if it's an existing `.xlsx`, else `discover_default_excel()`, else `filedialog.askopenfilename`. Cancelled picker → exit 0.
2. Call the same pipeline `__main__.main()` invokes (factor out the post-arg-parse body into a reusable function).
3. On success: `messagebox.showinfo("Done", f"Saved to {output}")`.
4. On exception: `messagebox.showerror(...)` with the message and log path, plus an "Open log folder" button that calls `os.startfile(log_dir)`.

### `src/logging_setup.py`

```python
def configure(log_path: Path | None, *, cli_mode: bool) -> None:
    """Attach handlers to the root logger."""
```

- File handler: `log_path` if provided, level `DEBUG`, includes timestamps and full tracebacks.
- Stream handler (`stdout`, level `INFO`): attached only when `cli_mode=True`.
- Log path is `<output>.log` — i.e., same stem as the output `.pptx` plus `.log`.
- If logging is needed before `output` is resolvable (early failure), fall back to `app_dir() / "sansa-excel2pptx-error.log"`.

### `__main__.py` flow changes

```python
def main() -> None:
    args = parser.parse_args()
    if is_frozen() and _no_flags_passed():
        sys.exit(gui.run())
    _run(args, cli_mode=True)

def _run(args, *, cli_mode: bool) -> None:
    # existing arg-resolution + pipeline logic
    # logging configured once output path is known
```

`_no_flags_passed()` checks `len(sys.argv) == 1` (double-click) or `len(sys.argv) == 2` with `argv[1]` an `.xlsx` (drag). Anything else means CLI mode.

## CI workflow (`.github/workflows/release.yml`)

```yaml
name: release
on:
  push:
    tags: ['v*']
jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.14' }
      - run: pip install -r requirements.txt pyinstaller
      - run: pyinstaller sansa-excel2pptx.spec
      - name: Smoke test
        run: |
          .\dist\sansa-excel2pptx.exe `
            --excel "excel\labalaba ginuma.xlsx" `
            --output smoke.pptx
          if (-not (Test-Path smoke.pptx)) { exit 1 }
      - name: Stage release
        run: |
          mkdir staging
          copy dist\sansa-excel2pptx.exe staging\
          copy "labalaba ginuma.pptx" staging\
          # README.txt generated inline
      - name: Zip
        run: |
          Compress-Archive -Path staging\* `
            -DestinationPath sansa-excel2pptx-${{ github.ref_name }}.zip
      - uses: softprops/action-gh-release@v2
        with:
          files: sansa-excel2pptx-${{ github.ref_name }}.zip
```

Zip layout (flat, no nested `dist/`):
```
sansa-excel2pptx.exe
labalaba ginuma.pptx
README.txt
```

## PyInstaller spec

`sansa-excel2pptx.spec` (checked in, used by both CI and local builds):

- `--onefile`
- `--windowed` (no always-on console; console attaches automatically when invoked from a parent terminal)
- `--name sansa-excel2pptx`
- Hidden imports pinned for `matplotlib` backends (PyInstaller usually finds them, but pinning avoids CI surprises)
- No bundled data files — template ships alongside the `.exe`.

## Testing

- **Smoke test in CI** — runs the freshly-built `.exe` against the repo's `excel/labalaba ginuma.xlsx`, asserts the output `.pptx` exists. Catches PyInstaller runtime breakage (missing hidden imports, wrong resource paths) before publish.
- **Local manual checks before tagging:** double-click flow with no co-located `.xlsx` (picker appears), with one `.xlsx` (auto-uses), drag-onto-icon, CLI flags. Forced error (e.g. delete the template) to verify the error dialog and log path.

## Out of scope

- macOS / Linux builds.
- Code signing (downloads will trigger SmartScreen warnings; user accepts the risk).
- Auto-update.
- A richer GUI with date picker / template chooser. (Section 3 question C — revisit if needed.)

## Open questions

None at design time.
