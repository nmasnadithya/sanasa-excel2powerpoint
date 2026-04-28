"""Apply the Sinhala (UN-Ganganee) font to a python-pptx run.

Sinhala is a complex script — Office picks the `<a:cs>` (complex script)
typeface, not `<a:latin>`. python-pptx exposes `.font.name` (which sets
`<a:latin>` only), so we patch the `<a:cs>` element directly via lxml.
"""
from __future__ import annotations

from lxml import etree
from pptx.text.text import _Run

from src.theme import LATIN_FONT, SINHALA_FONT


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def apply_sinhala_font(run: _Run, font_name: str = SINHALA_FONT,
                       latin_font: str = LATIN_FONT) -> None:
    """Set both `<a:latin>` and `<a:cs>` typefaces on a run.

    Without `<a:cs>`, PowerPoint falls back to a default font for Sinhala
    glyphs and the visual diverges from the reference deck.
    """
    run.font.name = latin_font
    rpr = run._r.get_or_add_rPr()
    cs = rpr.find(f"{{{_A_NS}}}cs")
    if cs is None:
        cs = etree.SubElement(rpr, f"{{{_A_NS}}}cs")
    cs.set("typeface", font_name)
