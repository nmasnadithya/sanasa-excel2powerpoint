"""Native PowerPoint pie chart with theme-coloured slices and readable legend."""
from __future__ import annotations

from typing import Sequence

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

from src import theme
from src.sinhala_font import apply_sinhala_to_font


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

    # Legend — readable size + black so it shows on the light theme.
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    legend_font = chart.legend.font
    legend_font.size = theme.CHART_LEGEND_SIZE
    legend_font.bold = False
    legend_font.color.rgb = theme.BLACK
    apply_sinhala_to_font(legend_font)

    # Per-slice fills + bold percentage data labels.
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_percentage = True
    dl.show_category_name = False
    dl.show_value = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = theme.CHART_DATA_LABEL_SIZE
    dl.font.bold = True
    dl.font.color.rgb = theme.BLACK
    dl.number_format = "0.0%"
    dl.number_format_is_linked = False

    series = plot.series[0]
    for idx, point in enumerate(series.points):
        fill = point.format.fill
        fill.solid()
        fill.fore_color.rgb = slice_colors[idx % len(slice_colors)]
