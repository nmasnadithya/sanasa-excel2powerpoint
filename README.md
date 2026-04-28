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
  theme.py                 # colors, fonts, dimensions calibrated to source slides 1-3
  sinhala_font.py          # XML helper for <a:cs> typeface (Sinhala)
  chart_writer.py          # native pie chart with theme-coloured slices
  builders/
    base.py                # Builder Protocol
    template_builder.py    # v1 builder (opens source deck, builds new slides, strips originals)
labalaba ginuma.pptx       # source deck — its slide-1 layout, master, theme are reused at runtime
excel/
  labalaba ginuma.xlsx     # input workbook
output/                    # generated decks (gitignored)
```

## How it works

1. `ExcelReader` opens the workbook, auto-detects the latest date column whose data isn't all zero, and exposes `column_for(sheet, date)`, `rows(sheet, range, col)` (with zero filtering), and `loan_surplus(date)` (label-based lookup).
2. `slide_specs.build_specs(reader, date)` returns a list of `SlideSpec` describing the deck. Slide 8 (loan surplus/deficit) is placed conditionally based on the sign of `loan_surplus`.
3. `TemplateBuilder` opens `labalaba ginuma.pptx`, snapshots the original 17 slide IDs, builds new slides on top using positions/fonts calibrated to source slides 1-3, then deletes the originals before saving. The cover slide uses the same layout as the source's slide 1, which carries the cover background image.

## Design notes

- **Source-calibrated geometry**: `theme.py` holds exact EMU coordinates lifted from source slides 1-3 (title position 457200×25398, table position 457200×1038134, chart position 457200×914400, etc). The intentional offset of slide 2's table — chosen by the user to compensate for projection alignment — is preserved.
- **Light theme everywhere**: slides 1-3 of the source set the intended palette (light bg, red `#C00000` titles, alternating-row tables, no header row). v1 standardises this across the whole deck — no dark-navy slides.
- **Summary slides at the end**: the summary table + chart pair for both income and expenses appears after all detail slides, sourced from the `සාරාංශය` summary sheet (stable cell positions even when line items are added).
- **Conditional slide 8**: when `බොල් හා අඩමාණ ණය` is negative, the slide appears at the end of the income block with title `අධි`; when positive, it appears between expenses and summary with title `ඌණ`. Zero suppresses the slide.
- **Zero-value filtering**: rows whose value is `0`, `None`, or empty are dropped from every table. Categories that filter to zero rows produce no slide.
- **Pagination**: detail tables paginate at 9 rows; pages 2+ get a `(අඛණ්ඩව)` suffix to maintain readability when projected.
- **Readable chart legends**: the source deck's chart legends are configured white-on-white (an artifact of an earlier theme); v1 forces black at 24 pt for readability while keeping `UN-Ganganee` as the typeface.

## Future work

- Hybrid / from-scratch builders selectable via `--builder` flag
- Additional chart types
- PyInstaller Windows packaging
- Visual regression testing via image diffs against a known-good run
- Cleaner orphan-part removal (the 17-slide strip leaves chart parts in the package, bloating the file to ~13 MB)
