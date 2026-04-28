# sansa-excel2powerpoint

Generates a Sinhala monthly financial PowerPoint deck from `excel/labalaba ginuma.xlsx`.

## Setup (one-time)

```bash
conda env create -f environment.yml
conda activate sansa-excel2powerpoint
```

The custom Sinhala font **UN-Ganganee** must be installed on the machine that *opens* the generated `.pptx`. Without it, Sinhala text falls back to a default font.

## Usage

```bash
# Latest non-empty month (auto-detected from the workbook):
python -m src

# Explicit month-end:
python -m src --date 2026-03-31

# Custom paths:
python -m src \
    --excel "excel/labalaba ginuma.xlsx" \
    --date 2026-02-28 \
    --output output/feb.pptx
```

By default the output goes to `output/<excel-stem>_<date>.pptx`.

## Project layout

```
src/
  __main__.py              # CLI entry
  excel_reader.py          # workbook access, date resolution, zero filtering
  slide_specs.py           # declarative deck plan
  theme.py                 # colors / fonts / sizes
  sinhala_font.py          # XML helper for <a:cs> typeface (Sinhala)
  chart_writer.py          # native pie chart with theme-coloured slices
  builders/
    base.py                # Builder Protocol
    template_builder.py    # v1 builder
templates/
  base.pptx                # slide-less template (master + theme only)
excel/
  labalaba ginuma.xlsx     # input workbook
output/                    # generated decks (gitignored)
```

## How it works

1. `ExcelReader` opens the workbook, auto-detects the latest date column whose data isn't all zero, and exposes:
   - `column_for(sheet, date)` → maps a date to its value column (`D` for income/expense sheets, `F`/`H`/etc for the summary sheet's paired columns)
   - `rows(sheet, range, col)` → reads (label, value) pairs, skipping zero-valued rows
   - `loan_surplus(date)` → label-based lookup of `බොල් හා අඩමාණ ණය` in the summary sheet
2. `slide_specs.build_specs(reader, date)` returns a list of `SlideSpec` describing the deck. Slide 8 (loan surplus/deficit) is placed conditionally based on the sign of `loan_surplus`.
3. `TemplateBuilder` opens `templates/base.pptx`, dispatches each spec to a layout-specific renderer, and saves to the output path. Tables paginate at 9 rows; pages 2+ get a `(අඛණ්ඩව)` suffix.

## Design notes

- **Light theme everywhere**: slides 1–3 of the original reference deck used the intended palette (light bg, red `#C00000` titles, alternating-row tables). v1 standardises this across the whole deck.
- **Summary slides at the end**: the summary table + chart pair for both income and expenses appear after all detail slides.
- **Conditional slide ordering**: when `බොල් හා අඩමාණ ණය` is negative, the slide-8 surplus appears at the end of the income block with title `අධි`; when positive, it appears between the expense and summary blocks with title `ඌණ`. Zero suppresses the slide.
- **Zero-value filtering**: rows whose value is `0`, `None`, or empty are dropped from every table. Categories that filter to zero rows produce no slide.
- **Summary slides source from `සාරාංශය`**: the summary block (income/expense overview tables and pie charts) reads from the summary sheet, which has stable cell positions even when line items are added to the income/expense detail sheets.

## Regenerating the template

`templates/base.pptx` is a stripped version of the original deck — same slide master, theme, and font registrations, but zero slides. To regenerate it from a freshly edited reference deck:

```python
from pptx import Presentation
prs = Presentation("labalaba ginuma.pptx")
ids = list(prs.slides._sldIdLst)
for sid in ids:
    rId = sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(sid)
prs.save("templates/base.pptx")
```

## Future work

- Cleaner orphan-part removal so `base.pptx` shrinks below 1 MB
- Hybrid / from-scratch builders selectable via `--builder` flag
- Additional chart types
- PyInstaller Windows packaging
