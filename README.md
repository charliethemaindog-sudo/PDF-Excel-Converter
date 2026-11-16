# PDF to Excel Converter

Pulls the tables out of PDF files and saves them as Excel (`.xlsx`) workbooks.

- Each table goes on its own sheet (`Page 1 Table 1`, `Page 1 Table 2`, ...), or all on one sheet with `--single-sheet`.
- Numbers become real numeric cells: `1,250.50` → 1250.5, `(4.00)` → -4, `$9` → 9, `15%` → 0.15.
- Pages with no table keep their text, one line per row, on a `Page N Text` sheet.
- Converts one file, several files or a whole folder in one go.

Works with PDFs that have selectable text. Scanned PDFs (pictures of pages) need OCR first.

## Install

Requires Python 3.9 or newer.

```bash
pip install -r requirements.txt
```

Or install the `pdf2excel` command itself:

```bash
pip install .
```

## Upload in your browser

```bash
python -m pdf2excel.web
```

Then open http://127.0.0.1:5000, drop in a PDF and click **Convert to Excel**. The `.xlsx` file downloads straight away. Press Ctrl+C in the terminal to stop the page.

The page runs only on your own computer, so your PDFs are never sent anywhere else.

## Command line

```bash
pdf2excel report.pdf                      # writes report.xlsx next to it
pdf2excel report.pdf -o tables.xlsx       # choose the output file
pdf2excel report.pdf -p 2-4               # only pages 2 to 4
pdf2excel report.pdf --single-sheet       # every table on one sheet
pdf2excel invoices/ -o converted/         # every PDF in a folder
```

Without installing, use `python -m pdf2excel` in place of `pdf2excel`.

| Option | What it does |
| --- | --- |
| `-o, --output` | Output `.xlsx` file, or a folder when converting several PDFs |
| `-p, --pages` | Pages to convert, e.g. `1,3-5` (default: all) |
| `-s, --single-sheet` | All tables on one sheet, each with a label row above it |
| `--keep-text` | Leave numbers as text |
| `--no-text-fallback` | Skip pages that have no tables |
| `--password` | Password for an encrypted PDF |

## Use from Python

```python
from pdf2excel import convert

path, count = convert("report.pdf", "report.xlsx", pages="1-3")
```

## Tests

```bash
pip install pytest reportlab
python -m pytest
```
