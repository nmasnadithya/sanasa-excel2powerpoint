"""v1 builder: opens templates/base.pptx (slide-less), adds slides per spec."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

from src import theme
from src.excel_reader import ExcelReader, INCOME_SHEET, EXPENSE_SHEET, SUMMARY_SHEET, Row
from src.sinhala_font import apply_sinhala_font
from src.slide_specs import SlideSpec, DataQuery
from src.chart_writer import build_pie_chart


MAX_ROWS_PER_TABLE_SLIDE = 9
CONTINUATION_SUFFIX = " (අඛණ්ඩව)"

DEFAULT_TEMPLATE = Path(__file__).resolve().parents[2] / "templates" / "base.pptx"


class TemplateBuilder:
    def __init__(self, template_path: Path | None = None):
        self.template_path = Path(template_path) if template_path else DEFAULT_TEMPLATE

    def build(
        self,
        specs: list[SlideSpec],
        reader: ExcelReader,
        target: date,
        output_path: Path,
    ) -> None:
        prs = Presentation(str(self.template_path))
        blank_layout = self._blank_layout(prs)

        for spec in specs:
            if spec.layout == "cover":
                self._render_cover(prs, blank_layout, spec)
            elif spec.layout == "table":
                self._render_table(prs, blank_layout, spec, reader, target)
            elif spec.layout == "big_number":
                self._render_big_number(prs, blank_layout, spec, reader, target)
            elif spec.layout == "chart":
                self._render_chart(prs, blank_layout, spec, reader, target)
            else:
                raise ValueError(f"Unknown layout {spec.layout!r}")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(output_path))

    @staticmethod
    def _blank_layout(prs):
        for layout in prs.slide_layouts:
            if layout.name == "Blank":
                return layout
        return prs.slide_layouts[-1]

    # --- Renderers ---------------------------------------------------------

    def _render_cover(self, prs, layout, spec: SlideSpec) -> None:
        slide = prs.slides.add_slide(layout)
        self._set_white_background(slide)

        slide_w, slide_h = prs.slide_width, prs.slide_height
        title_box = slide.shapes.add_textbox(
            Inches(0.5),
            Emu(slide_h // 2 - Inches(1.5).emu),
            slide_w - Inches(1).emu,
            Inches(1.5),
        )
        self._fill_text(title_box, spec.title, size=Pt(48), color=theme.RED, bold=True)

        if spec.subtitle:
            subtitle_box = slide.shapes.add_textbox(
                Inches(0.5),
                Emu(slide_h // 2 + Inches(0.2).emu),
                slide_w - Inches(1).emu,
                Inches(1.5),
            )
            self._fill_text(
                subtitle_box, spec.subtitle, size=theme.SUBTITLE_SIZE, color=theme.RED
            )

    def _render_table(
        self, prs, layout, spec: SlideSpec, reader: ExcelReader, target: date
    ) -> None:
        rows = self._fetch_rows(spec.data, reader, target)
        if not rows:
            return  # zero-only category — skip slide entirely

        # Paginate
        chunks: list[list[Row]] = [
            rows[i : i + MAX_ROWS_PER_TABLE_SLIDE]
            for i in range(0, len(rows), MAX_ROWS_PER_TABLE_SLIDE)
        ]
        for page_idx, chunk in enumerate(chunks):
            slide = prs.slides.add_slide(layout)
            self._set_white_background(slide)
            title = spec.title if page_idx == 0 else spec.title + CONTINUATION_SUFFIX
            self._draw_title(prs, slide, title)
            self._draw_table(prs, slide, chunk)

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
        self._draw_title(prs, slide, spec.title)

        # Center the formatted number
        slide_w, slide_h = prs.slide_width, prs.slide_height
        num_box = slide.shapes.add_textbox(
            Inches(0.5),
            Emu(slide_h // 2 - Inches(1.5).emu),
            slide_w - Inches(1).emu,
            Inches(2.0),
        )
        self._fill_text(
            num_box,
            self._fmt_number(value),
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
        self._draw_title(prs, slide, spec.title)

        slide_w, slide_h = prs.slide_width, prs.slide_height
        chart_w = Inches(10)
        chart_h = Inches(5.5)
        x = Emu((slide_w - chart_w.emu) // 2)
        y = Inches(1.5)
        build_pie_chart(
            slide,
            x,
            y,
            chart_w,
            chart_h,
            labels=[r.label for r in rows],
            values=[r.value for r in rows],
            slice_colors=theme.PIE_PALETTE,
        )

    # --- Helpers -----------------------------------------------------------

    def _fetch_rows(
        self, query: DataQuery | None, reader: ExcelReader, target: date
    ) -> list[Row]:
        if query is None:
            return []
        col = reader.column_for(query.sheet, target)
        return reader.rows(query.sheet, query.rows, col, query.label_col)

    def _draw_title(self, prs, slide, text: str) -> None:
        slide_w = prs.slide_width
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.3), slide_w - Inches(1).emu, Inches(0.9)
        )
        self._fill_text(
            title_box, text, size=theme.TITLE_SIZE, color=theme.RED, bold=True
        )

    def _draw_table(self, prs, slide, rows: list[Row]) -> None:
        slide_w, slide_h = prs.slide_width, prs.slide_height
        n_rows = len(rows)
        table_w = slide_w - Inches(1).emu
        table_h = min(Inches(5.5).emu, n_rows * Inches(0.55).emu)

        x = Inches(0.5)
        y = Inches(1.4)
        shape = slide.shapes.add_table(n_rows, 2, x, y, table_w, table_h)
        table = shape.table

        # column widths: 70% label, 30% value
        table.columns[0].width = Emu(int(table_w * 0.70))
        table.columns[1].width = Emu(table_w - int(table_w * 0.70))

        for i, row in enumerate(rows):
            self._populate_cell(
                table.cell(i, 0),
                row.label,
                bg=theme.TABLE_ROW_A if i % 2 == 0 else theme.TABLE_ROW_B,
                align=PP_ALIGN.LEFT,
            )
            self._populate_cell(
                table.cell(i, 1),
                self._fmt_number(row.value),
                bg=theme.TABLE_ROW_A if i % 2 == 0 else theme.TABLE_ROW_B,
                align=PP_ALIGN.RIGHT,
                apply_sinhala=False,
            )

    def _populate_cell(
        self,
        cell,
        text: str,
        bg,
        align: PP_ALIGN,
        apply_sinhala: bool = True,
    ) -> None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = True
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = text
        run.font.size = theme.TABLE_TEXT_SIZE
        run.font.color.rgb = theme.BLACK
        if apply_sinhala:
            apply_sinhala_font(run)
        else:
            run.font.name = theme.LATIN_FONT

    def _fill_text(
        self,
        textbox,
        text: str,
        size,
        color,
        bold: bool = False,
        align: PP_ALIGN = PP_ALIGN.LEFT,
        apply_sinhala: bool = True,
    ) -> None:
        tf = textbox.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = text
        run.font.size = size
        run.font.color.rgb = color
        run.font.bold = bold
        if apply_sinhala:
            apply_sinhala_font(run)
        else:
            run.font.name = theme.LATIN_FONT

    @staticmethod
    def _set_white_background(slide) -> None:
        # Slides inherit master/layout bg; explicit white guards against the
        # original deck's dark navy bleeding through if any layout has dark fill.
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme.LIGHT_BG

    @staticmethod
    def _fmt_number(value: float) -> str:
        return f"{value:,.2f}"
