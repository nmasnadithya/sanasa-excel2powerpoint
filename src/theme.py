"""Color, font, and layout constants — calibrated to slides 1-3 of the source deck."""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt


# Colors --------------------------------------------------------------------
RED = RGBColor(0xC0, 0x00, 0x00)            # title color on slides 2+
RED_BRIGHT = RGBColor(0xFF, 0x00, 0x00)     # subtitle on slide 1 (the cover)
BLACK = RGBColor(0x02, 0x00, 0x18)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

LIGHT_BG = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_A = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_B = RGBColor(0xEC, 0xEC, 0xEC)

# Fonts ---------------------------------------------------------------------
SINHALA_FONT = "UN-Ganganee"
LATIN_FONT = "Segoe UI"

# Text sizes (matched to source slides 1-3) ---------------------------------
COVER_SUBTITLE_SIZE = Pt(54)                # slide 1: 685800 EMU
TABLE_TITLE_SIZE = Pt(48)                   # slide 2 title: 609600 EMU
CHART_TITLE_SIZE = Pt(40)                   # slide 3 title: 508000 EMU
TABLE_CELL_SIZE = Pt(34)                    # slide 2 cells: 431800 EMU
BIG_NUMBER_SIZE = Pt(80)
CHART_LEGEND_SIZE = Pt(32)                  # readable from a hall; source had 36pt
CHART_DATA_LABEL_SIZE = Pt(28)              # bold percentage labels

# Geometry (slides 2 and 3 — exact positions from the source deck) ----------
TABLE_TITLE_LEFT = Emu(457200)
TABLE_TITLE_TOP = Emu(25398)
TABLE_TITLE_WIDTH = Emu(11277600)
TABLE_TITLE_HEIGHT = Emu(830997)

TABLE_LEFT = Emu(457200)
TABLE_TOP = Emu(1038134)
TABLE_WIDTH = Emu(9716530)
# Wider value column than the source so 12-character formatted numbers
# (e.g. "1,234,426.00") render on a single line.
TABLE_COL0_WIDTH = Emu(6216530)             # label column ~64%
TABLE_COL1_WIDTH = Emu(3500000)             # value column ~36%
TABLE_ROW_HEIGHT = Emu(564457)

CHART_TITLE_LEFT = Emu(457200)
CHART_TITLE_TOP = Emu(76200)
CHART_TITLE_WIDTH = Emu(11277600)
CHART_TITLE_HEIGHT = Emu(685800)

CHART_LEFT = Emu(457200)
CHART_TOP = Emu(914400)
CHART_WIDTH = Emu(11277600)
CHART_HEIGHT = Emu(5486400)

# Cover slide placeholder overrides (slide 1 — the dark cover) --------------
COVER_TITLE_LEFT = Emu(762000)
COVER_TITLE_TOP = Emu(1651000)
COVER_TITLE_WIDTH = Emu(10668000)
COVER_TITLE_HEIGHT = Emu(1524000)

COVER_SUBTITLE_LEFT = Emu(1524000)
COVER_SUBTITLE_TOP = Emu(3619500)
COVER_SUBTITLE_WIDTH = Emu(9364394)
COVER_SUBTITLE_HEIGHT = Emu(2236762)

COVER_DECOR_LEFT = Emu(4826000)             # the small accent rectangle
COVER_DECOR_TOP = Emu(3403600)
COVER_DECOR_WIDTH = Emu(2540000)
COVER_DECOR_HEIGHT = Emu(38100)
COVER_DECOR_COLOR = RGBColor(0x4A, 0x90, 0xD9)

# Pie palette ---------------------------------------------------------------
PIE_PALETTE = [
    RGBColor(0xC0, 0x00, 0x00),
    RGBColor(0xE8, 0xB3, 0x0E),
    RGBColor(0x4A, 0x90, 0xD9),
    RGBColor(0x6A, 0xA8, 0x4F),
    RGBColor(0x96, 0x4B, 0x00),
    RGBColor(0x6F, 0x42, 0xC1),
    RGBColor(0xD9, 0x53, 0x4F),
    RGBColor(0x20, 0x71, 0x9E),
]
