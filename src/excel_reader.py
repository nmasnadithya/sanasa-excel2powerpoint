"""Read and resolve data from the Sansa monthly financial workbook."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook


INCOME_SHEET = "ආදායම්"
EXPENSE_SHEET = "වියදම්"
SUMMARY_SHEET = "සාරාංශය"

LOAN_SURPLUS_LABEL = "බොල් හා අඩමාණ ණය"


@dataclass(frozen=True)
class Row:
    label: str
    value: float


def _normalize(text: object) -> str:
    if text is None:
        return ""
    return str(text).replace("\u200d", "").strip()


def _parse_date_cell(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _is_zero(value: object) -> bool:
    if value is None or value == "":
        return True
    try:
        return float(value) == 0.0
    except (TypeError, ValueError):
        return False


class ExcelReader:
    def __init__(self, path: Path | str):
        self.path = Path(path)
        self.wb: Workbook = load_workbook(self.path, data_only=True)

    def _sheet(self, name: str):
        return self.wb[name]

    def date_columns(self, sheet: str = INCOME_SHEET) -> list[tuple[date, str]]:
        """Return (date, column-letter) pairs for every date header in row 1.

        Skips the first column (label) and the totals column (`මුළු එකතුව`).
        Used by both detail sheets (one column per date) and summary sheet
        (paired columns — only the value column is returned).
        """
        ws = self._sheet(sheet)
        results: list[tuple[date, str]] = []
        for col_idx in range(2, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx).value
            d = _parse_date_cell(cell)
            if d is not None:
                results.append((d, get_column_letter(col_idx)))
        return results

    def available_dates(self, sheet: str = INCOME_SHEET) -> list[date]:
        return [d for d, _ in self.date_columns(sheet)]

    def column_for(self, sheet: str, target: date) -> str:
        for d, col in self.date_columns(sheet):
            if d == target:
                return col
        raise ValueError(
            f"Date {target.isoformat()} not found in sheet {sheet!r}. "
            f"Available: {[d.isoformat() for d in self.available_dates(sheet)]}"
        )

    def latest_populated_date(self) -> date:
        """Pick the most recent date whose data column has at least one
        non-zero, non-None value in the income or expense sheet.

        Falls back to earlier months if the most recent column is empty
        (e.g. when a fresh month-end column has been added but not yet filled).
        """
        income_cols = self.date_columns(INCOME_SHEET)
        if not income_cols:
            raise ValueError(f"No date headers found in sheet {INCOME_SHEET!r}")
        income_cols.sort(key=lambda pair: pair[0], reverse=True)
        for d, col in income_cols:
            if self._has_data(INCOME_SHEET, col) or self._has_data(EXPENSE_SHEET, col):
                return d
        # Nothing populated anywhere — fall back to the latest header
        return income_cols[0][0]

    def _has_data(self, sheet: str, col_letter: str) -> bool:
        ws = self._sheet(sheet)
        col_idx = ws[col_letter + "1"].column
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if not _is_zero(value):
                return True
        return False

    def rows(
        self,
        sheet: str,
        row_range: tuple[int, int] | tuple[int, ...],
        value_col: str,
        label_col: str = "A",
    ) -> list[Row]:
        """Return non-zero rows from a range or explicit row list.

        Filters out rows whose value cell is None / "" / 0. Both top-level
        category rows and detail line items use this method.
        """
        ws = self._sheet(sheet)
        if len(row_range) == 2 and isinstance(row_range[0], int) and row_range[0] < row_range[1]:
            indices: list[int] = list(range(row_range[0], row_range[1] + 1))
        else:
            indices = list(row_range)

        out: list[Row] = []
        for r in indices:
            label = _normalize(ws[f"{label_col}{r}"].value)
            value = ws[f"{value_col}{r}"].value
            if not label or _is_zero(value):
                continue
            out.append(Row(label=label, value=float(value)))
        return out

    def loan_surplus(self, target: date) -> float:
        """Look up `බොල් හා අඩමාණ ණය` in the summary sheet by label.

        Returns the signed value at the target date's value column.
        Zero (or missing) means no slide-8 should be emitted.
        """
        ws = self._sheet(SUMMARY_SHEET)
        col = self.column_for(SUMMARY_SHEET, target)
        target_label = _normalize(LOAN_SURPLUS_LABEL)
        for row_idx in range(2, ws.max_row + 1):
            if _normalize(ws.cell(row=row_idx, column=1).value) == target_label:
                value = ws[f"{col}{row_idx}"].value
                return float(value) if value is not None else 0.0
        raise ValueError(
            f"Could not find row labelled {LOAN_SURPLUS_LABEL!r} in {SUMMARY_SHEET!r}"
        )
