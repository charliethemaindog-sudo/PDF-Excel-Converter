import pytest
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

from pdf2excel.cli import main
from pdf2excel.converter import convert, parse_pages, to_number

ROWS = [
    ["Item", "Qty", "Price"],
    ["Apples", "3", "1,250.50"],
    ["Pears", "10", "(4.00)"],
]


@pytest.fixture
def sample_pdf(tmp_path):
    path = tmp_path / "sample.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=letter)
    table = Table(ROWS)
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 1, "black")]))
    styles = getSampleStyleSheet()
    doc.build([table, PageBreak(), Paragraph("Just some notes, no table here.", styles["Normal"])])
    return path


def test_to_number():
    assert to_number("1,234") == 1234
    assert to_number("1,250.50") == 1250.5
    assert to_number("(4.00)") == -4.0
    assert to_number("$9") == 9
    assert to_number("15%") == pytest.approx(0.15)
    assert to_number("abc") == "abc"
    assert to_number("2024-01-01") == "2024-01-01"


def test_parse_pages():
    assert parse_pages(None, 3) == [0, 1, 2]
    assert parse_pages("1,3", 3) == [0, 2]
    assert parse_pages("2-", 3) == [1, 2]
    with pytest.raises(ValueError):
        parse_pages("5", 3)


def test_convert_sheet_per_table(sample_pdf, tmp_path):
    out, count = convert(sample_pdf, tmp_path / "out.xlsx")
    assert count == 2
    wb = load_workbook(out)
    assert wb.sheetnames == ["Page 1 Table 1", "Page 2 Text"]
    values = [list(r) for r in wb["Page 1 Table 1"].iter_rows(values_only=True)]
    assert values == [["Item", "Qty", "Price"], ["Apples", 3, 1250.5], ["Pears", 10, -4.0]]
    assert "notes" in wb["Page 2 Text"]["A1"].value


def test_convert_single_sheet_and_no_fallback(sample_pdf, tmp_path):
    out, count = convert(sample_pdf, tmp_path / "one.xlsx", single_sheet=True, text_fallback=False)
    assert count == 1
    wb = load_workbook(out)
    assert wb.sheetnames == ["All Tables"]
    assert wb.active["A1"].value == "Page 1 - Table 1"
    assert wb.active["A2"].value == "Item"


def test_keep_text(sample_pdf, tmp_path):
    out, _ = convert(sample_pdf, tmp_path / "t.xlsx", convert_numbers=False, pages="1")
    assert load_workbook(out)["Page 1 Table 1"]["B2"].value == "3"


def test_cli_folder(sample_pdf, tmp_path):
    out_dir = tmp_path / "out"
    assert main([str(sample_pdf.parent), "-o", str(out_dir)]) == 0
    assert (out_dir / "sample.xlsx").exists()


def test_cli_missing_file(tmp_path):
    assert main([str(tmp_path / "nope.pdf")]) == 1
