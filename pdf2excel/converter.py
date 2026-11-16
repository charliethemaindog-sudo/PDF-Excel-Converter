"""Core PDF -> Excel conversion logic."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

_NUMBER_RE = re.compile(r"^\(?-?[$€£]?\s?[\d,]*\.?\d+\)?%?$")
_HEADER_FONT = Font(bold=True)
_HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
_MAX_SHEET_NAME = 31


@dataclass
class Table:
    page: int
    index: int  # 1-based position of the table on its page
    rows: list[list[object]] = field(default_factory=list)
    is_text: bool = False  # True when this is a plain-text fallback, not a detected table


def parse_pages(spec: str | None, page_count: int) -> list[int]:
    """Turn '1,3-5' into zero-based page indexes. None means all pages."""
    if not spec:
        return list(range(page_count))
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s) if start_s else 1
            end = int(end_s) if end_s else page_count
        else:
            start = end = int(part)
        if start < 1 or end > page_count or start > end:
            raise ValueError(f"Page range '{part}' is outside 1-{page_count}")
        pages.extend(p - 1 for p in range(start, end + 1) if p - 1 not in pages)
    return pages


def to_number(value: object) -> object:
    """Convert numeric-looking text such as '1,234.50', '(42)', '$9', '15%' to numbers."""
    if not isinstance(value, str):
        return value
    text = value.strip()
    if not text or not _NUMBER_RE.match(text):
        return text
    negative = text.startswith("(") and text.endswith(")")
    percent = text.endswith("%")
    cleaned = re.sub(r"[()$€£,%\s]", "", text)
    try:
        number = float(cleaned)
    except ValueError:
        return text
    if negative:
        number = -number
    if percent:
        number /= 100
    return int(number) if number.is_integer() and not percent and "." not in cleaned else number


def _clean_cell(cell: object, convert_numbers: bool) -> object:
    if cell is None:
        return None
    text = " ".join(str(cell).split())  # collapse line breaks inside cells
    return to_number(text) if convert_numbers else text


def extract_tables(
    pdf_path: str | Path,
    pages: str | None = None,
    convert_numbers: bool = True,
    text_fallback: bool = True,
    password: str | None = None,
) -> list[Table]:
    """Extract every table from the PDF. Pages with no tables yield their text lines."""
    tables: list[Table] = []
    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page_idx in parse_pages(pages, len(pdf.pages)):
            page = pdf.pages[page_idx]
            found = page.extract_tables()
            found = [t for t in found if any(any(c not in (None, "") for c in row) for row in t)]
            if found:
                for i, raw in enumerate(found, start=1):
                    rows = [[_clean_cell(c, convert_numbers) for c in row] for row in raw]
                    tables.append(Table(page=page_idx + 1, index=i, rows=rows))
            elif text_fallback:
                text = page.extract_text() or ""
                lines = [[line] for line in text.splitlines() if line.strip()]
                if lines:
                    tables.append(Table(page=page_idx + 1, index=1, rows=lines, is_text=True))
    return tables


def _sheet_name(table: Table, used: set[str]) -> str:
    base = f"Page {table.page} Text" if table.is_text else f"Page {table.page} Table {table.index}"
    name = base[:_MAX_SHEET_NAME]
    n = 2
    while name in used:
        suffix = f" ({n})"
        name = base[: _MAX_SHEET_NAME - len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


def _write_rows(ws, rows: list[list[object]], start_row: int, bold_first: bool) -> int:
    for r, row in enumerate(rows):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=start_row + r, column=c, value=value)
            if bold_first and r == 0:
                cell.font = _HEADER_FONT
                cell.fill = _HEADER_FILL
    return start_row + len(rows)


def _autosize(ws) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                widths[cell.column] = max(widths.get(cell.column, 0), len(str(cell.value)))
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 2, 8), 60)


def convert(
    pdf_path: str | Path,
    output_path: str | Path | None = None,
    pages: str | None = None,
    single_sheet: bool = False,
    convert_numbers: bool = True,
    text_fallback: bool = True,
    password: str | None = None,
) -> tuple[Path, int]:
    """Convert a PDF to .xlsx. Returns (output path, number of tables/text blocks written)."""
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    output = Path(output_path) if output_path else pdf_path.with_suffix(".xlsx")

    tables = extract_tables(pdf_path, pages, convert_numbers, text_fallback, password)

    wb = Workbook()
    if not tables:
        ws = wb.active
        ws.title = "No tables found"
        ws["A1"] = "No tables or text could be extracted. The PDF may be a scanned image."
    elif single_sheet:
        ws = wb.active
        ws.title = "All Tables"
        row = 1
        for table in tables:
            label = f"Page {table.page} - " + ("Text" if table.is_text else f"Table {table.index}")
            ws.cell(row=row, column=1, value=label).font = Font(bold=True, italic=True)
            row = _write_rows(ws, table.rows, row + 1, bold_first=not table.is_text) + 1
        _autosize(ws)
    else:
        wb.remove(wb.active)
        used: set[str] = set()
        for table in tables:
            ws = wb.create_sheet(_sheet_name(table, used))
            _write_rows(ws, table.rows, 1, bold_first=not table.is_text)
            if not table.is_text:
                ws.freeze_panes = "A2"
            _autosize(ws)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output, len(tables)
