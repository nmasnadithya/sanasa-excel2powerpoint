"""Apply the Sinhala (UN-Ganganee) font to a python-pptx run or font wrapper.

Sinhala is a complex script — Office picks the `<a:cs>` (complex script)
typeface, not `<a:latin>`. python-pptx exposes `.font.name` (which sets
`<a:latin>` only), so we patch the `<a:cs>` element directly via lxml.
"""
from __future__ import annotations

from lxml import etree

from src.theme import LATIN_FONT, SINHALA_FONT


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def apply_sinhala_font(run, font_name: str = SINHALA_FONT,
                       latin_font: str = LATIN_FONT) -> None:
    """Set both `<a:latin>` and `<a:cs>` typefaces on a run."""
    run.font.name = latin_font
    rpr = run._r.get_or_add_rPr()
    _ensure_cs(rpr, font_name)


def apply_sinhala_to_font(font, font_name: str = SINHALA_FONT,
                          latin_font: str = LATIN_FONT) -> None:
    """Apply Sinhala typeface to a python-pptx Font (e.g. chart legend font).

    Charts expose Font wrappers without runs — we reach into the underlying
    rPr / defRPr XML element. Without `<a:cs>`, Sinhala falls back to a
    default font and the chart legend becomes unreadable.
    """
    font.name = latin_font
    rpr = _font_rpr(font)
    if rpr is not None:
        _ensure_cs(rpr, font_name)


def _font_rpr(font):
    """Return the rPr-like element backing a python-pptx Font wrapper."""
    # Most Font wrappers expose `_rPr` (text run) or `_defRPr` (paragraph default
    # / chart text properties). Both behave like rPr for our purposes.
    for attr in ("_rPr", "_defRPr"):
        el = getattr(font, attr, None)
        if el is not None:
            return el
    # Some wrappers expose the parent element via different names; fall back
    # to scanning attributes for any lxml element.
    for attr in dir(font):
        if attr.startswith("_") and not attr.startswith("__"):
            try:
                el = getattr(font, attr)
            except Exception:
                continue
            if hasattr(el, "tag") and hasattr(el, "set"):
                return el
    return None


def _ensure_cs(rpr, font_name: str) -> None:
    cs = rpr.find(f"{{{_A_NS}}}cs")
    if cs is None:
        cs = etree.SubElement(rpr, f"{{{_A_NS}}}cs")
    cs.set("typeface", font_name)
