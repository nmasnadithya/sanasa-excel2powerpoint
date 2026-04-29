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
    try:
        run_pipeline(excel, args.date, template, args.output, cli_mode=True)
    except Exception:
        # In --windowed PyInstaller builds, Python's default exception display
        # routes to a Tk dialog that can't render in headless CI. Force the
        # traceback to stderr so failures are visible in logs.
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
