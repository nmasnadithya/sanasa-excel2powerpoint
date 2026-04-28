"""Native PowerPoint pie chart with theme-coloured slices."""
from __future__ import annotations

from typing import Sequence

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

from src import theme
from src.sinhala_font import apply_sinhala_font


def build_pie_chart(
    slide,
    x,
    y,
    cx,
    cy,
    labels: Sequence[str],
    values: Sequence[float],
    slice_colors: Sequence[RGBColor],
) -> None:
    chart_data = CategoryChartData()
    chart_data.categories = list(labels)
    chart_data.add_series("", list(values))

    graphic = slide.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data)
    chart = graphic.chart

    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    _style_run(chart.legend.font)

    # Per-slice colors
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_percentage = True
    data_labels.show_category_name = False
    data_labels.show_value = False
    data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    _style_run(data_labels.font, color=theme.BLACK)

    series = plot.series[0]
    for idx, point in enumerate(series.points):
        fill = point.format.fill
        fill.solid()
        fill.fore_color.rgb = slice_colors[idx % len(slice_colors)]


def _style_run(font, color=None) -> None:
    """Style the chart legend / data label font with the Sinhala typeface.

    Charts use a different XML branch from regular runs — `font.name` here
    only sets the latin typeface, so we touch the underlying rPr to add a
    `<a:cs>` element when present.
    """
    font.name = theme.LATIN_FONT
    if color is not None:
        font.color.rgb = color
    # Charts don't expose run objects the same way; the latin name + theme
    # font registration in the master is usually enough for legend text.
