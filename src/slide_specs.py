"""Declarative deck plan: build_specs(reader, target_date) → list[SlideSpec]."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from src.excel_reader import (
    ExcelReader,
    INCOME_SHEET,
    EXPENSE_SHEET,
    SUMMARY_SHEET,
)


COVER_TITLE = "සී/ර බිබිලාදෙනිය සණස සමිතිය"
COVER_SUBTITLE_TEMPLATE = "{date} මාසය සදහා ආදායම් වියදම් ඇස්තමේන්තුව"

LOAN_SURPLUS_TITLE_NEGATIVE = "බොල් හා අඩමාණ ණය අධි වෙන් කිරීම"
LOAN_SURPLUS_TITLE_POSITIVE = "බොල් හා අඩමාණ ණය ඌණ වෙන් කිරීම"


@dataclass(frozen=True)
class DataQuery:
    sheet: str
    rows: tuple[int, int] | tuple[int, ...]
    label_col: str = "A"


@dataclass(frozen=True)
class SlideSpec:
    layout: str  # "cover" | "table" | "chart" | "big_number"
    title: str
    subtitle: Optional[str] = None
    data: Optional[DataQuery] = None
    computed_key: Optional[str] = None  # e.g. "loan_surplus"


def build_specs(reader: ExcelReader, target_date: date) -> list[SlideSpec]:
    surplus = reader.loan_surplus(target_date)
    formatted_date = target_date.strftime("%Y/%m/%d")

    cover = SlideSpec(
        layout="cover",
        title=COVER_TITLE,
        subtitle=COVER_SUBTITLE_TEMPLATE.format(date=formatted_date),
    )

    income_details = [
        SlideSpec(
            layout="table",
            title="පොළී ආදායම",
            data=DataQuery(INCOME_SHEET, (3, 13)),
        ),
        SlideSpec(
            layout="table",
            title="බැංකු පොළී ආදායම්",
            data=DataQuery(INCOME_SHEET, (15, 19)),
        ),
        SlideSpec(
            layout="table",
            title="වෙනත් ආදායම්",
            data=DataQuery(INCOME_SHEET, (21, 25)),
        ),
        SlideSpec(
            layout="table",
            title="ලැබිය යුතු බැංකු පොළී",
            data=DataQuery(INCOME_SHEET, (27, 30)),
        ),
    ]

    expense_details = [
        SlideSpec(
            layout="table",
            title="ගෙවූ පොළී",
            data=DataQuery(EXPENSE_SHEET, (3, 14)),
        ),
        SlideSpec(
            layout="table",
            title="වෙනත් වියදම්",
            data=DataQuery(EXPENSE_SHEET, (16, 43)),
        ),
        SlideSpec(
            layout="table",
            title="ක්ෂය වෙන් කිරීම",
            data=DataQuery(EXPENSE_SHEET, (48, 52)),
        ),
        SlideSpec(
            layout="table",
            title="ගෙවිය යුතු පොළී",
            data=DataQuery(EXPENSE_SHEET, (45, 46)),
        ),
    ]

    summary_block = [
        SlideSpec(
            layout="table",
            title="ආදායම",
            data=DataQuery(SUMMARY_SHEET, (4, 7)),
        ),
        SlideSpec(
            layout="chart",
            title="ආදායම් බෙදීම",
            data=DataQuery(SUMMARY_SHEET, (4, 7)),
        ),
        SlideSpec(
            layout="table",
            title="වියදම",
            data=DataQuery(SUMMARY_SHEET, (11, 14)),
        ),
        SlideSpec(
            layout="chart",
            title="වියදම් බෙදීම",
            data=DataQuery(SUMMARY_SHEET, (11, 14)),
        ),
    ]

    specs: list[SlideSpec] = [cover, *income_details]
    if surplus < 0:
        specs.append(_loan_surplus_spec(LOAN_SURPLUS_TITLE_NEGATIVE))
    specs.extend(expense_details)
    if surplus > 0:
        specs.append(_loan_surplus_spec(LOAN_SURPLUS_TITLE_POSITIVE))
    specs.extend(summary_block)
    return specs


def _loan_surplus_spec(title: str) -> SlideSpec:
    return SlideSpec(
        layout="big_number",
        title=title,
        computed_key="loan_surplus",
    )
