"""CLI entry point: python -m src --excel ... --date ... --output ..."""
from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

from src.excel_reader import ExcelReader
from src.slide_specs import build_specs
from src.builders.template_builder import TemplateBuilder, DEFAULT_TEMPLATE


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EXCEL = REPO_ROOT / "excel" / "labalaba ginuma.xlsx"


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="sansa-excel2pptx",
        description="Generate a Sinhala financial PPTX from a monthly Excel workbook.",
    )
    parser.add_argument("--excel", type=Path, default=DEFAULT_EXCEL,
                        help="Path to the Excel workbook")
    parser.add_argument("--date", type=_parse_date, default=None,
                        help="Target month-end date (YYYY-MM-DD). "
                             "Defaults to the latest date with non-empty data.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE,
                        help="Path to templates/base.pptx")
    parser.add_argument("--output", type=Path, default=None,
                        help="Output .pptx path (default: output/<excel>_<date>.pptx)")
    args = parser.parse_args()

    reader = ExcelReader(args.excel)
    target = args.date or reader.latest_populated_date()

    if args.output is None:
        out_dir = REPO_ROOT / "output"
        stem = args.excel.stem
        args.output = out_dir / f"{stem}_{target.isoformat()}.pptx"

    specs = build_specs(reader, target)
    print(f"Generating {len(specs)} spec(s) for {target.isoformat()} → {args.output}")
    builder = TemplateBuilder(template_path=args.template)
    builder.build(specs, reader, target, args.output)
    print(f"✓ Wrote {args.output}")


if __name__ == "__main__":
    main()
