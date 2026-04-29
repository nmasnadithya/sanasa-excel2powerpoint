"""Matplotlib-rendered chart images for slides that python-pptx can't do well.

Image kinds dispatched here:
    - heatmap_income / heatmap_expense
    - waterfall
    - sankey
    - kpi_tiles
    - small_multiples_income / small_multiples_expense

Sinhala font registration runs once at module import. matplotlib's font cache
is process-local and doesn't auto-discover system fonts on macOS/Windows.
"""
from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path

import logging

import matplotlib

matplotlib.use("Agg")  # headless
# Silence the "Font family 'Iskoola Pota' not found" stream — that font is
# the Windows fallback in our chain, harmless to skip on macOS.
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Rectangle
from matplotlib.sankey import Sankey

from src import theme
from src.excel_reader import (
    ExcelReader, INCOME_SHEET, EXPENSE_SHEET, SUMMARY_SHEET,
)


class NoDataError(Exception):
    """Raised when a chart has no data to render — caller should skip the slide."""


# --- Sinhala font registration ---------------------------------------------

def _register_sinhala_fonts() -> str | None:
    """Try platform paths; return the first successfully registered family."""
    candidates: list[str] = []
    if sys.platform == "darwin":
        home = os.path.expanduser("~")
        candidates = [
            f"{home}/Library/Fonts/UN-Ganganee.ttf",
            "/Library/Fonts/UN-Ganganee.ttf",
            "/System/Library/Fonts/Supplemental/Sinhala MN.ttc",
            "/System/Library/Fonts/Supplemental/Sinhala Sangam MN.ttc",
        ]
    elif sys.platform == "win32":
        candidates = [
            r"C:\Windows\Fonts\UN-Ganganee.ttf",
            os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts\UN-Ganganee.ttf"),
            r"C:\Windows\Fonts\Iskoola.ttf",
        ]
    else:  # linux/other
        candidates = [
            "/usr/share/fonts/truetype/UN-Ganganee.ttf",
        ]

    registered: list[str] = []
    for p in candidates:
        if os.path.exists(p):
            try:
                font_manager.fontManager.addfont(p)
                registered.append(p)
            except Exception as e:
                print(f"[image_writer] could not register {p}: {e}", file=sys.stderr)

    plt.rcParams["font.family"] = theme.MATPLOTLIB_FONT_FAMILY
    plt.rcParams["axes.unicode_minus"] = False

    if registered:
        return registered[0]
    print("[image_writer] WARNING: no Sinhala fonts registered — labels may "
          "appear as tofu boxes", file=sys.stderr)
    return None


_FONT_REGISTERED = _register_sinhala_fonts()


# --- Public API ------------------------------------------------------------

FIGSIZE = (13.33, 5.5)   # matches CHART_WIDTH/HEIGHT 16:9 chart-area
DPI = 200


def render(kind: str, reader: ExcelReader, target: date,
           output_path: Path) -> None:
    if kind == "heatmap_income":
        _render_heatmap(reader, INCOME_SHEET, (3, 30), output_path)
    elif kind == "heatmap_expense":
        _render_heatmap(reader, EXPENSE_SHEET, (3, 52), output_path)
    elif kind == "waterfall":
        _render_waterfall(reader, target, output_path)
    elif kind == "sankey":
        _render_sankey(reader, target, output_path)
    elif kind == "kpi_tiles":
        _render_kpi_tiles(reader, target, output_path)
    elif kind == "small_multiples_income":
        _render_small_multiples(reader, SUMMARY_SHEET, (4, 7), output_path)
    elif kind == "small_multiples_expense":
        _render_small_multiples(reader, SUMMARY_SHEET, (11, 14), output_path)
    else:
        raise ValueError(f"Unknown image kind {kind!r}")


# --- Helpers ---------------------------------------------------------------

def _hex(rgb) -> str:
    return "#" + str(rgb)


def _palette_hex() -> list[str]:
    return [_hex(c) for c in theme.PIE_PALETTE]


def _fmt_month(d: date) -> str:
    months_si = ["", "ජන", "පෙබ", "මාර්", "අප්‍ර", "මැයි", "ජූනි",
                 "ජූලි", "අගෝ", "සැප්", "ඔක්", "නොවැ", "දෙසැ"]
    return f"{months_si[d.month]} {d.year % 100:02d}"


def _save(fig, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=DPI, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# --- Heatmap ----------------------------------------------------------------

def _render_heatmap(reader: ExcelReader, sheet: str,
                    row_range: tuple[int, int], output_path: Path) -> None:
    multi = reader.rows_multi(sheet, row_range)
    dates = reader.populated_dates()
    # Filter rows that have at least one populated value
    rows = [mr for mr in multi if any(d in mr.values for d in dates)]
    if not rows or not dates:
        raise NoDataError(f"No data for heatmap on {sheet!r}")

    # Sort rows by total magnitude — biggest at top
    rows.sort(key=lambda mr: sum(mr.values.values()), reverse=True)

    labels = [mr.label for mr in rows]
    matrix = [[mr.values.get(d, 0.0) for d in dates] for mr in rows]

    fig, ax = plt.subplots(figsize=FIGSIZE)
    im = ax.imshow(matrix, aspect="auto", cmap="Reds")

    ax.set_xticks(range(len(dates)))
    ax.set_xticklabels([_fmt_month(d) for d in dates], fontsize=14)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=10)

    # Annotate each cell with its value (only for non-zero)
    for i in range(len(labels)):
        for j in range(len(dates)):
            v = matrix[i][j]
            if v > 0:
                ax.text(j, i, f"{v / 1000:.0f}K", ha="center", va="center",
                        fontsize=8, color="black")

    cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    cbar.ax.tick_params(labelsize=12)
    fig.tight_layout()
    _save(fig, output_path)


# --- Waterfall --------------------------------------------------------------

def _render_waterfall(reader: ExcelReader, target: date,
                      output_path: Path) -> None:
    income = reader.adjusted_monthly_totals(INCOME_SHEET).get(target, 0.0)
    expense_categories = reader.rows_multi(SUMMARY_SHEET, (11, 14))
    surplus = reader.loan_surplus(target)

    if income == 0 or not expense_categories:
        raise NoDataError("No data for waterfall")

    labels = ["ආදායම"]
    values = [income]
    for mr in expense_categories:
        v = mr.values.get(target, 0.0)
        if v > 0:
            labels.append(mr.label)
            values.append(-v)
    if surplus > 0:
        labels.append("බොල් හා අඩමාණ ණය ඌණ වෙන් කිරීම")
        values.append(-surplus)
    labels.append("ශුද්ධ ලාභය")
    net = income - sum(-v for v in values[1:] if v < 0)
    values.append(net)

    fig, ax = plt.subplots(figsize=FIGSIZE)
    cumulative = 0.0
    bar_colors: list[str] = []
    for i, v in enumerate(values):
        if i == 0 or i == len(values) - 1:
            bar_colors.append(_hex(theme.LINE_PROFIT_COLOR))
            ax.bar(i, abs(v), bottom=0, color=bar_colors[-1])
            cumulative = v if i == 0 else cumulative
        elif v < 0:
            bar_colors.append(_hex(theme.LINE_EXPENSE_COLOR))
            ax.bar(i, abs(v), bottom=cumulative + v, color=bar_colors[-1])
            cumulative += v

    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=20, ha="right", fontsize=11)
    ax.set_ylabel("රු.", fontsize=12)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    fig.tight_layout()
    _save(fig, output_path)


# --- Sankey -----------------------------------------------------------------

def _render_sankey(reader: ExcelReader, target: date,
                   output_path: Path) -> None:
    income_cats = reader.rows_multi(SUMMARY_SHEET, (4, 7))
    expense_cats = reader.rows_multi(SUMMARY_SHEET, (11, 14))
    surplus = reader.loan_surplus(target)

    inflows: list[tuple[str, float]] = []
    for mr in income_cats:
        v = mr.values.get(target, 0.0)
        if v > 0:
            inflows.append((mr.label, v))
    if surplus < 0:
        inflows.append(("බොල් හා අඩමාණ ණය අධි වෙන් කිරීම", abs(surplus)))

    outflows: list[tuple[str, float]] = []
    for mr in expense_cats:
        v = mr.values.get(target, 0.0)
        if v > 0:
            outflows.append((mr.label, v))
    if surplus > 0:
        outflows.append(("බොල් හා අඩමාණ ණය ඌණ වෙන් කිරීම", surplus))

    if not inflows or not outflows:
        raise NoDataError("Sankey needs both inflows and outflows")

    total_in = sum(v for _, v in inflows)
    total_out = sum(v for _, v in outflows)
    net = total_in - total_out
    if net > 0:
        outflows.append(("ශුද්ධ ලාභය", net))
    elif net < 0:
        inflows.append(("ශුද්ධ අලාභය", -net))

    flows = [v for _, v in inflows] + [-v for _, v in outflows]
    labels = [n for n, _ in inflows] + [n for n, _ in outflows]
    orientations = [0] * len(inflows) + [0] * len(outflows)

    fig, ax = plt.subplots(figsize=FIGSIZE)
    ax.axis("off")
    sankey = Sankey(ax=ax, scale=1.0 / max(flows), gap=0.5, format="%.0f")
    sankey.add(flows=flows, labels=labels, orientations=orientations,
               trunklength=2.0)
    diagrams = sankey.finish()
    for d in diagrams:
        for t in d.texts:
            t.set_fontsize(10)
    fig.tight_layout()
    _save(fig, output_path)


# --- KPI tiles --------------------------------------------------------------

def _render_kpi_tiles(reader: ExcelReader, target: date,
                      output_path: Path) -> None:
    dates = reader.populated_dates()
    if target not in dates:
        raise NoDataError(f"target {target} not populated")
    income_t = reader.adjusted_monthly_totals(INCOME_SHEET).get(target, 0.0)
    expense_t = reader.adjusted_monthly_totals(EXPENSE_SHEET).get(target, 0.0)
    net_t = income_t - expense_t
    savings_rate = (net_t / income_t * 100) if income_t else 0

    prev_dates = [d for d in dates if d < target]
    if prev_dates:
        prev = max(prev_dates)
        income_p = reader.adjusted_monthly_totals(INCOME_SHEET).get(prev, 0.0)
        expense_p = reader.adjusted_monthly_totals(EXPENSE_SHEET).get(prev, 0.0)
        income_growth = ((income_t - income_p) / income_p * 100) if income_p else 0
        expense_growth = ((expense_t - expense_p) / expense_p * 100) if expense_p else 0
    else:
        income_growth = None
        expense_growth = None

    tiles = [
        ("ශුද්ධ ලාභය", f"{net_t:,.0f}", _hex(theme.LINE_PROFIT_COLOR)),
        ("ඉතුරුම් අනුපාතය", f"{savings_rate:.1f}%",
         _hex(theme.LINE_INCOME_COLOR if savings_rate >= 0
              else theme.LINE_EXPENSE_COLOR)),
        ("ආදායම් වර්ධනය",
         "—" if income_growth is None else f"{income_growth:+.1f}%",
         _hex(theme.LINE_INCOME_COLOR if (income_growth or 0) >= 0
              else theme.LINE_EXPENSE_COLOR)),
        ("වියදම් වර්ධනය",
         "—" if expense_growth is None else f"{expense_growth:+.1f}%",
         _hex(theme.LINE_EXPENSE_COLOR if (expense_growth or 0) >= 0
              else theme.LINE_INCOME_COLOR)),
    ]

    fig, axes = plt.subplots(1, 4, figsize=FIGSIZE)
    for ax, (label, value, color) in zip(axes, tiles):
        ax.axis("off")
        ax.add_patch(Rectangle((0, 0), 1, 1, facecolor=color, alpha=0.12,
                               edgecolor=color, linewidth=2))
        ax.text(0.5, 0.65, value, ha="center", va="center",
                fontsize=28, fontweight="bold", color=color)
        ax.text(0.5, 0.25, label, ha="center", va="center", fontsize=14)
    fig.tight_layout()
    _save(fig, output_path)


# --- Small multiples --------------------------------------------------------

def _render_small_multiples(reader: ExcelReader, sheet: str,
                            row_range: tuple[int, int],
                            output_path: Path) -> None:
    dates = reader.populated_dates()
    multi = reader.rows_multi(sheet, row_range)
    if not multi or not dates:
        raise NoDataError("No data for small multiples")

    n = len(multi)
    cols = 2 if n > 2 else n
    rows_count = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows_count, cols, figsize=FIGSIZE,
                             sharex=True)
    axes = axes.flatten() if n > 1 else [axes]

    palette = _palette_hex()
    for ax, mr, color in zip(axes, multi, palette):
        ys = [mr.values.get(d, 0.0) for d in dates]
        xs = [_fmt_month(d) for d in dates]
        ax.plot(xs, ys, color=color, marker="o", linewidth=2.5)
        ax.set_title(mr.label, fontsize=12)
        ax.tick_params(axis="x", labelsize=10)
        ax.tick_params(axis="y", labelsize=9)
        ax.grid(True, linestyle="--", alpha=0.3)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    # Hide unused axes
    for ax in axes[len(multi):]:
        ax.axis("off")
    fig.tight_layout()
    _save(fig, output_path)
