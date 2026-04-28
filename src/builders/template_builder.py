"""v1 builder — opens the source deck, uses its layouts/master, builds new
slides with positions and fonts calibrated to slides 1-3 of the source."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches

from src import theme
from src.excel_reader import ExcelReader, Row
from src.sinhala_font import apply_sinhala_font
from src.slide_specs import SlideSpec, DataQuery
from src.chart_writer import build_pie_chart


MAX_ROWS_PER_TABLE_SLIDE = 9
CONTINUATION_SUFFIX = " (අඛණ්ඩව)"

DEFAULT_SOURCE = (
    Path(__file__).resolve().parents[2] / "labalaba ginuma.pptx"
)

_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class TemplateBuilder:
    """Open the source deck, build new slides on top, then strip originals."""

    def __init__(self, source_path: Path | None = None):
        self.source_path = Path(source_path) if source_path else DEFAULT_SOURCE

    # --- Public API --------------------------------------------------------

    def build(
        self,
        specs: list[SlideSpec],
        reader: ExcelReader,
        target: date,
        output_path: Path,
    ) -> None:
        prs = Presentation(str(self.source_path))
        cover_layout = prs.slides[0].slide_layout
        blank_layout = self._find_layout(prs, "Blank")

        original_sld_ids = list(prs.slides._sldIdLst)

        for spec in specs:
            if spec.layout == "cover":
                self._render_cover(prs, cover_layout, spec)
            elif spec.layout == "table":
                self._render_table(prs, blank_layout, spec, reader, target)
            elif spec.layout == "big_number":
                self._render_big_number(prs, blank_layout, spec, reader, target)
            elif spec.layout == "chart":
                self._render_chart(prs, blank_layout, spec, reader, target)
            else:
                raise ValueError(f"Unknown layout {spec.layout!r}")

        self._remove_slides(prs, original_sld_ids)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(output_path))

    # --- Renderers ---------------------------------------------------------

    def _render_cover(self, prs, layout, spec: SlideSpec) -> None:
        slide = prs.slides.add_slide(layout)
        # Cover keeps its layout-defined background image — don't paint over it.

        title_ph, subtitle_ph = self._cover_placeholders(slide)

        if title_ph is not None:
            title_ph.left = theme.COVER_TITLE_LEFT
            title_ph.top = theme.COVER_TITLE_TOP
            title_ph.width = theme.COVER_TITLE_WIDTH
            title_ph.height = theme.COVER_TITLE_HEIGHT
            self._set_textframe_text(
                title_ph.text_frame,
                spec.title,
                size=None,                # let layout/auto-size decide
                color=None,
                bold=True,
                align=None,
            )

        if subtitle_ph is not None and spec.subtitle:
            subtitle_ph.left = theme.COVER_SUBTITLE_LEFT
            subtitle_ph.top = theme.COVER_SUBTITLE_TOP
            subtitle_ph.width = theme.COVER_SUBTITLE_WIDTH
            subtitle_ph.height = theme.COVER_SUBTITLE_HEIGHT
            self._set_textframe_text(
                subtitle_ph.text_frame,
                spec.subtitle,
                size=theme.COVER_SUBTITLE_SIZE,
                color=theme.RED_BRIGHT,
                bold=True,                # user wants bold on the date subtitle
                align=None,
            )

        decor = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            theme.COVER_DECOR_LEFT,
            theme.COVER_DECOR_TOP,
            theme.COVER_DECOR_WIDTH,
            theme.COVER_DECOR_HEIGHT,
        )
        decor.fill.solid()
        decor.fill.fore_color.rgb = theme.COVER_DECOR_COLOR
        decor.line.fill.background()

    def _render_table(
        self, prs, layout, spec: SlideSpec, reader: ExcelReader, target: date
    ) -> None:
        rows = self._fetch_rows(spec.data, reader, target)
        if not rows:
            return

        chunks = _distribute_evenly(rows, MAX_ROWS_PER_TABLE_SLIDE)
        for page_idx, chunk in enumerate(chunks):
            slide = prs.slides.add_slide(layout)
            self._set_white_background(slide)
            title = spec.title if page_idx == 0 else spec.title + CONTINUATION_SUFFIX
            self._draw_table_title(slide, title)
            self._draw_table(slide, chunk)

    def _render_big_number(
        self, prs, layout, spec: SlideSpec, reader: ExcelReader, target: date
    ) -> None:
        if spec.computed_key != "loan_surplus":
            raise ValueError(f"Unknown computed_key {spec.computed_key!r}")
        value = abs(reader.loan_surplus(target))
        if value == 0:
            return

        slide = prs.slides.add_slide(layout)
        self._set_white_background(slide)
        self._draw_table_title(slide, spec.title)

        # Position the number in the lower-middle band of the slide so it
        # reads as the focal element, not crammed against the title.
        slide_w = prs.slide_width
        slide_h = prs.slide_height
        box_h = Emu(2_500_000)
        num_box = slide.shapes.add_textbox(
            Emu(0),
            Emu(int(slide_h * 0.45) - box_h // 2),
            slide_w,
            box_h,
        )
        self._set_textframe_text(
            num_box.text_frame,
            f"{value:,.2f}",
            size=theme.BIG_NUMBER_SIZE,
            color=theme.BLACK,
            bold=True,
            align=PP_ALIGN.CENTER,
            apply_sinhala=False,
        )

    def _render_chart(
        self, prs, layout, spec: SlideSpec, reader: ExcelReader, target: date
    ) -> None:
        rows = self._fetch_rows(spec.data, reader, target)
        if not rows:
            return

        slide = prs.slides.add_slide(layout)
        self._set_white_background(slide)
        self._draw_chart_title(slide, spec.title)
        build_pie_chart(
            slide,
            x=theme.CHART_LEFT,
            y=theme.CHART_TOP,
            cx=theme.CHART_WIDTH,
            cy=theme.CHART_HEIGHT,
            labels=[r.label for r in rows],
            values=[r.value for r in rows],
            slice_colors=theme.PIE_PALETTE,
        )

    # --- Helpers -----------------------------------------------------------

    def _draw_table_title(self, slide, text: str) -> None:
        tb = slide.shapes.add_textbox(
            theme.TABLE_TITLE_LEFT,
            theme.TABLE_TITLE_TOP,
            theme.TABLE_TITLE_WIDTH,
            theme.TABLE_TITLE_HEIGHT,
        )
        self._set_textframe_text(
            tb.text_frame,
            text,
            size=theme.TABLE_TITLE_SIZE,
            color=theme.RED,
            bold=False,
            align=PP_ALIGN.CENTER,
        )

    def _draw_chart_title(self, slide, text: str) -> None:
        tb = slide.shapes.add_textbox(
            theme.CHART_TITLE_LEFT,
            theme.CHART_TITLE_TOP,
            theme.CHART_TITLE_WIDTH,
            theme.CHART_TITLE_HEIGHT,
        )
        self._set_textframe_text(
            tb.text_frame,
            text,
            size=theme.CHART_TITLE_SIZE,
            color=theme.RED,
            bold=False,
            align=PP_ALIGN.CENTER,
        )

    def _draw_table(self, slide, rows: list[Row]) -> None:
        n_rows = len(rows)
        table_h = theme.TABLE_ROW_HEIGHT * n_rows
        shape = slide.shapes.add_table(
            n_rows,
            2,
            theme.TABLE_LEFT,
            theme.TABLE_TOP,
            theme.TABLE_WIDTH,
            table_h,
        )
        table = shape.table

        table.first_row = False
        table.first_col = False
        table.horz_banding = False
        table.vert_banding = False

        table.columns[0].width = theme.TABLE_COL0_WIDTH
        table.columns[1].width = theme.TABLE_COL1_WIDTH

        for i, row in enumerate(rows):
            table.rows[i].height = theme.TABLE_ROW_HEIGHT
            bg = theme.TABLE_ROW_A if i % 2 == 0 else theme.TABLE_ROW_B
            self._populate_cell(
                table.cell(i, 0), row.label, bg=bg,
                align=PP_ALIGN.LEFT, bold=False,
                apply_sinhala=True, word_wrap=True,
            )
            self._populate_cell(
                table.cell(i, 1), f"{row.value:,.2f}", bg=bg,
                align=PP_ALIGN.RIGHT, bold=True,
                apply_sinhala=False, word_wrap=False,
            )

    def _populate_cell(self, cell, text: str, bg, align, bold: bool,
                       apply_sinhala: bool, word_wrap: bool) -> None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.12)
        cell.margin_right = Inches(0.12)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = word_wrap
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = text
        run.font.size = theme.TABLE_CELL_SIZE
        run.font.color.rgb = theme.BLACK
        run.font.bold = bold
        if apply_sinhala:
            apply_sinhala_font(run)
        else:
            run.font.name = theme.LATIN_FONT

    def _set_textframe_text(self, tf, text: str, size, color, bold: bool,
                            align, apply_sinhala: bool = True) -> None:
        tf.clear()
        tf.word_wrap = True
        para = tf.paragraphs[0]
        if align is not None:
            para.alignment = align
        run = para.add_run()
        run.text = text
        if size is not None:
            run.font.size = size
        if color is not None:
            run.font.color.rgb = color
        run.font.bold = bold
        if apply_sinhala:
            apply_sinhala_font(run)
        else:
            run.font.name = theme.LATIN_FONT

    def _fetch_rows(self, query: DataQuery | None, reader: ExcelReader,
                    target: date) -> list[Row]:
        if query is None:
            return []
        col = reader.column_for(query.sheet, target)
        return reader.rows(query.sheet, query.rows, col, query.label_col)

    @staticmethod
    def _set_white_background(slide) -> None:
        """Override the master's dark navy background with explicit white.

        The source deck's master is `#1B2A4A`; only slides 1-3 override it.
        Without this, every new content slide inherits the dark navy.
        """
        cSld = slide.element.find(f"{{{_P_NS}}}cSld")
        existing_bg = cSld.find(f"{{{_P_NS}}}bg")
        if existing_bg is not None:
            cSld.remove(existing_bg)
        bg = etree.SubElement(cSld, f"{{{_P_NS}}}bg")
        bg_pr = etree.SubElement(bg, f"{{{_P_NS}}}bgPr")
        solid = etree.SubElement(bg_pr, f"{{{_A_NS}}}solidFill")
        etree.SubElement(solid, f"{{{_A_NS}}}srgbClr", val="FFFFFF")
        etree.SubElement(bg_pr, f"{{{_A_NS}}}effectLst")
        # `bg` must come BEFORE `spTree` per ECMA-376; lxml SubElement
        # appends to end. Re-order if needed.
        sp_tree = cSld.find(f"{{{_P_NS}}}spTree")
        if sp_tree is not None and bg.getnext() is not sp_tree:
            cSld.remove(bg)
            cSld.insert(list(cSld).index(sp_tree), bg)

    @staticmethod
    def _cover_placeholders(slide):
        title_ph = None
        subtitle_ph = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 0:
                title_ph = ph
            elif ph.placeholder_format.idx == 1:
                subtitle_ph = ph
        return title_ph, subtitle_ph

    @staticmethod
    def _find_layout(prs, name: str):
        for layout in prs.slide_layouts:
            if layout.name == name:
                return layout
        return prs.slide_layouts[-1]

    @staticmethod
    def _remove_slides(prs, sld_ids) -> None:
        sldIdLst = prs.slides._sldIdLst
        for sid in sld_ids:
            rId = sid.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            prs.part.drop_rel(rId)
            sldIdLst.remove(sid)


def _distribute_evenly(items, max_per_slide: int) -> list[list]:
    """Split into pages with no slide exceeding `max_per_slide`, distributing
    remainders so page sizes differ by at most 1.

    Per the user's spec: when items > max_per_slide,
    n_slides = (items // max_per_slide) + 1, then balance.
    For 22 items at max 9: 3 slides of 8, 7, 7.
    """
    n = len(items)
    if n <= max_per_slide:
        return [list(items)]
    n_slides = (n // max_per_slide) + 1
    base = n // n_slides
    extra = n % n_slides
    chunks: list[list] = []
    idx = 0
    for i in range(n_slides):
        size = base + (1 if i < extra else 0)
        chunks.append(list(items[idx:idx + size]))
        idx += size
    return chunks
