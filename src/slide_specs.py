"""Declarative deck plan: build_specs(reader, target_date) → list[SlideSpec]."""
from __future__ import annotations

from dataclasses import dataclass
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

INCOME_TOTAL_LABEL = "මුළු ආදායම"
EXPENSE_TOTAL_LABEL = "මුළු වියදම"


@dataclass(frozen=True)
class DataQuery:
    sheet: str
    rows: tuple[int, int] | tuple[int, ...]
    label_col: str = "A"


@dataclass(frozen=True)
class SlideSpec:
    layout: str
    # "cover" | "table" | "chart" | "big_number"
    # | "bar_compare" | "line_trend" | "stacked_bar"
    # | "ytd_table" | "delta_table" | "top_n_table" | "image"
    title: str
    subtitle: Optional[str] = None
    data: Optional[DataQuery] = None
    computed_key: Optional[str] = None
    # surplus_role: "income" | "expense"; when set on table/chart specs the
    # renderer injects the loan-surplus row/slice when its sign matches
    surplus_role: Optional[str] = None
    # image_kind: dispatch key for image_writer.render(...)
    image_kind: Optional[str] = None
    # compare_kind: "income" | "expense"; controls the green/red Δ% color rule
    compare_kind: Optional[str] = None
    # top_n: number of items for top_n_table
    top_n: Optional[int] = None


def build_specs(reader: ExcelReader, target_date: date) -> list[SlideSpec]:
    surplus = reader.loan_surplus(target_date)
    formatted_date = target_date.strftime("%Y/%m/%d")

    cover = SlideSpec(
        layout="cover",
        title=COVER_TITLE,
        subtitle=COVER_SUBTITLE_TEMPLATE.format(date=formatted_date),
    )

    income_details = [
        SlideSpec(layout="table", title="පොළී ආදායම",
                  data=DataQuery(INCOME_SHEET, (3, 13))),
        SlideSpec(layout="table", title="බැංකු පොළී ආදායම්",
                  data=DataQuery(INCOME_SHEET, (15, 19))),
        SlideSpec(layout="table", title="වෙනත් ආදායම්",
                  data=DataQuery(INCOME_SHEET, (21, 25))),
        SlideSpec(layout="table", title="ලැබිය යුතු බැංකු පොළී",
                  data=DataQuery(INCOME_SHEET, (27, 30))),
    ]

    expense_details = [
        SlideSpec(layout="table", title="ගෙවූ පොළී",
                  data=DataQuery(EXPENSE_SHEET, (3, 14))),
        SlideSpec(layout="table", title="වෙනත් වියදම්",
                  data=DataQuery(EXPENSE_SHEET, (16, 43))),
        SlideSpec(layout="table", title="ක්ෂය වෙන් කිරීම",
                  data=DataQuery(EXPENSE_SHEET, (48, 52))),
        SlideSpec(layout="table", title="ගෙවිය යුතු පොළී",
                  data=DataQuery(EXPENSE_SHEET, (45, 46))),
    ]

    # Summary block — surplus_role triggers conditional injection of the
    # surplus row/slice + the adjusted-total row.
    summary_block = [
        SlideSpec(
            layout="table",
            title="ආදායම",
            data=DataQuery(SUMMARY_SHEET, (4, 7)),
            surplus_role="income",
        ),
        SlideSpec(
            layout="chart",
            title="ආදායම් බෙදීම",
            data=DataQuery(SUMMARY_SHEET, (4, 7)),
            surplus_role="income",
        ),
        SlideSpec(
            layout="table",
            title="වියදම",
            data=DataQuery(SUMMARY_SHEET, (11, 14)),
            surplus_role="expense",
        ),
        SlideSpec(
            layout="chart",
            title="වියදම් බෙදීම",
            data=DataQuery(SUMMARY_SHEET, (11, 14)),
            surplus_role="expense",
        ),
    ]

    # v2 trends block — appended after summary
    trends_block = [
        SlideSpec(layout="bar_compare", title="මාසික ආදායම සහ වියදම"),
        SlideSpec(layout="line_trend", title="ශුද්ධ ලාභය",
                  computed_key="net_profit"),
        SlideSpec(layout="line_trend", title="සමුච්චිත ආදායම සහ වියදම",
                  computed_key="cumulative_io"),
        SlideSpec(layout="stacked_bar", title="මාසික ආදායම් සංයුතිය",
                  data=DataQuery(SUMMARY_SHEET, (4, 7)),
                  surplus_role="income"),
        SlideSpec(layout="stacked_bar", title="මාසික වියදම් සංයුතිය",
                  data=DataQuery(SUMMARY_SHEET, (11, 14)),
                  surplus_role="expense"),
        SlideSpec(layout="ytd_table", title="වසරට සාරාංශය — ආදායම",
                  data=DataQuery(SUMMARY_SHEET, (4, 7))),
        SlideSpec(layout="ytd_table", title="වසරට සාරාංශය — වියදම්",
                  data=DataQuery(SUMMARY_SHEET, (11, 14))),
        SlideSpec(layout="delta_table", title="පෙර මාසයට වඩා වෙනස — ආදායම්",
                  data=DataQuery(SUMMARY_SHEET, (4, 7)),
                  compare_kind="income"),
        SlideSpec(layout="delta_table", title="පෙර මාසයට වඩා වෙනස — වියදම්",
                  data=DataQuery(SUMMARY_SHEET, (11, 14)),
                  compare_kind="expense"),
        SlideSpec(layout="top_n_table", title="මෙම මාසයේ ලොකුම ආදායම්",
                  data=DataQuery(INCOME_SHEET, (2, 30)), top_n=5),
        SlideSpec(layout="top_n_table", title="මෙම මාසයේ ලොකුම වියදම්",
                  data=DataQuery(EXPENSE_SHEET, (2, 52)), top_n=5),
        SlideSpec(layout="image", title="ආදායම් වර්ග × මාස",
                  image_kind="heatmap_income"),
        SlideSpec(layout="image", title="වියදම් වර්ග × මාස",
                  image_kind="heatmap_expense"),
        SlideSpec(layout="image", title="ලාභ රටාව",
                  image_kind="waterfall"),
        SlideSpec(layout="image", title="ආදායම් / වියදම් ප්‍රවාහය",
                  image_kind="sankey"),
        SlideSpec(layout="image", title="ප්‍රමුඛ දර්ශක",
                  image_kind="kpi_tiles"),
        SlideSpec(layout="image", title="ආදායම් වර්ග ප්‍රවණතා",
                  image_kind="small_multiples_income"),
        SlideSpec(layout="image", title="වියදම් වර්ග ප්‍රවණතා",
                  image_kind="small_multiples_expense"),
    ]

    specs: list[SlideSpec] = [cover, *income_details]
    if surplus < 0:
        specs.append(_loan_surplus_spec(LOAN_SURPLUS_TITLE_NEGATIVE))
    specs.extend(expense_details)
    if surplus > 0:
        specs.append(_loan_surplus_spec(LOAN_SURPLUS_TITLE_POSITIVE))
    specs.extend(summary_block)
    specs.extend(trends_block)
    return specs


def _loan_surplus_spec(title: str) -> SlideSpec:
    return SlideSpec(
        layout="big_number",
        title=title,
        computed_key="loan_surplus",
    )
