import io

import pytest
from openpyxl import load_workbook

from pdf2excel.text_parser import parse_text
from pdf2excel.web import app


@pytest.mark.parametrize("text, mode, expected", [
    ("Name\tQty\nApples\t3\nPears\t10", "tab", [["Name", "Qty"], ["Apples", 3], ["Pears", 10]]),
    ('Name,Amount\n"Smith, John","1,200"\nJane,50', "comma",
     [["Name", "Amount"], ["Smith, John", 1200], ["Jane", 50]]),
    ("| Item | Cost |\n|---|---|\n| Rent | $1,200 |", "pipe", [["Item", "Cost"], ["Rent", 1200]]),
    ("a;b\n1;2", "semicolon", [["a", "b"], [1, 2]]),
    ("Date        Description      Amount\n01/02/2024  Coffee shop      4.50\n01/03/2024  Rent payment  1,200.00",
     "spaces", [["Date", "Description", "Amount"], ["01/02/2024", "Coffee shop", 4.5],
                ["01/03/2024", "Rent payment", 1200.0]]),
    ("01/02 Coffee shop 4.50 995.50\n01/03 Rent payment 1,200.00 (204.50)", "numbers",
     [["01/02", "Coffee shop", 4.5, 995.5], ["01/03", "Rent payment", 1200.0, -204.5]]),
    ("Name: Jane Doe\nTotal: $1,234.56", "label", [["Name", "Jane Doe"], ["Total", 1234.56]]),
])
def test_auto_detect(text, mode, expected):
    result = parse_text(text)
    assert result["mode"] == mode
    assert result["rows"] == expected


def test_header_detection():
    assert parse_text("Item\tCost\nRent\t1200")["header"] is True
    assert parse_text("Name: Jane\nAge: 30")["header"] is False
    assert parse_text("1\t2\n3\t4")["header"] is False


def test_ragged_rows_are_padded():
    assert parse_text("a\tb\tc\n1\t2")["rows"] == [["a", "b", "c"], [1, 2, ""]]


def test_forced_delimiter_and_keep_text():
    result = parse_text("x,y\t1", delimiter="tab", convert_numbers=False)
    assert result["rows"] == [["x,y", "1"]]


def test_blank_and_plain_text():
    assert parse_text("   \n\n")["rows"] == []
    assert parse_text("just one line of prose")["rows"] == [["just one line of prose"]]


def test_convert_text_endpoint():
    client = app.test_client()
    res = client.post("/convert-text", json={"text": "Item\tCost\nRent\t1,200.50"})
    assert res.status_code == 200
    ws = load_workbook(io.BytesIO(res.data)).active
    assert ws.title == "Pasted Table"
    assert ws["A1"].font.bold and ws["B2"].value == 1200.5


def test_parse_text_endpoint_and_errors():
    client = app.test_client()
    data = client.post("/parse-text", json={"text": "a,b\n1,2"}).get_json()
    assert data["mode"] == "comma" and data["rows"][1] == [1, 2]
    assert client.post("/convert-text", json={"text": ""}).status_code == 400
    assert client.post("/parse-text", json={"text": "a", "delimiter": "nope"}).status_code == 400
