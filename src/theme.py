"""Color and font constants — the visual theme for v1.

Colors picked to match slides 1–3 of the reference deck (light bg, red title,
alternating-row tables). All slides in v1 use this single palette.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.util import Pt


RED = RGBColor(0xC0, 0x00, 0x00)
BLACK = RGBColor(0x02, 0x00, 0x18)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

LIGHT_BG = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_A = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_B = RGBColor(0xF2, 0xF2, 0xF2)
TABLE_HEADER_BG = RGBColor(0xC0, 0x00, 0x00)

SINHALA_FONT = "UN-Ganganee"
LATIN_FONT = "Segoe UI"

TITLE_SIZE = Pt(36)
SUBTITLE_SIZE = Pt(28)
TABLE_TEXT_SIZE = Pt(20)
BIG_NUMBER_SIZE = Pt(96)

# Six-color palette for pie-chart slices. Coherent with the red/light theme
# while staying distinguishable when projected.
PIE_PALETTE = [
    RGBColor(0xC0, 0x00, 0x00),  # red
    RGBColor(0xE8, 0xB3, 0x0E),  # gold
    RGBColor(0x4A, 0x90, 0xD9),  # blue
    RGBColor(0x6A, 0xA8, 0x4F),  # green
    RGBColor(0x96, 0x4B, 0x00),  # brown
    RGBColor(0x6F, 0x42, 0xC1),  # purple
    RGBColor(0xD9, 0x53, 0x4F),  # coral
    RGBColor(0x20, 0x71, 0x9E),  # navy-blue
]
