import io

from openpyxl import load_workbook

from pdf2excel.web import app
from test_converter import sample_pdf  # noqa: F401  (pytest fixture)


def client():
    app.config["TESTING"] = True
    return app.test_client()


def test_index_page():
    res = client().get("/")
    assert res.status_code == 200
    assert b"PDF to Excel" in res.data


def test_upload_returns_excel(sample_pdf):  # noqa: F811
    res = client().post("/convert", data={"pdf": (open(sample_pdf, "rb"), "sample.pdf")},
                        content_type="multipart/form-data")
    assert res.status_code == 200
    assert res.headers["X-Table-Count"] == "2"
    assert "sample.xlsx" in res.headers["Content-Disposition"]
    wb = load_workbook(io.BytesIO(res.data))
    assert wb["Page 1 Table 1"]["C2"].value == 1250.5


def test_rejects_non_pdf():
    res = client().post("/convert", data={"pdf": (io.BytesIO(b"hi"), "notes.txt")},
                        content_type="multipart/form-data")
    assert res.status_code == 400


def test_bad_page_range(sample_pdf):  # noqa: F811
    res = client().post("/convert", data={"pdf": (open(sample_pdf, "rb"), "s.pdf"), "pages": "9"},
                        content_type="multipart/form-data")
    assert res.status_code == 400
    assert "outside" in res.get_json()["error"]


def test_corrupt_pdf():
    res = client().post("/convert", data={"pdf": (io.BytesIO(b"not a pdf"), "bad.pdf")},
                        content_type="multipart/form-data")
    assert res.status_code == 400
