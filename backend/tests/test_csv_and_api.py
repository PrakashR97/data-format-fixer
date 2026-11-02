from io import BytesIO

from openpyxl import Workbook, load_workbook


def test_csv_semicolon_and_cp1252(run, make_csv):
    p = make_csv("name;qty;price\nCafé;1;2.5\nTea;2;3.75\n", encoding="cp1252")
    report, wb = run(p, kind="csv")
    ws = wb.worksheets[0]
    assert ws["A2"].value == "Café"
    assert ws["B2"].value == 1 and ws["C3"].value == 3.75
    assert report["csv"]["delimiter"] == "semicolon"
    assert "cannot store" in report["csv"]["note"]


def test_csv_tab_delimiter_and_bom(run, make_csv):
    p = make_csv("﻿a\tb\n1\t2\n", encoding="utf-8")
    report, wb = run(p, kind="csv")
    assert wb.worksheets[0]["B2"].value == 2
    assert report["csv"]["delimiter"] == "tab"


def test_csv_formula_like_text_stays_literal(run, make_csv):
    p = make_csv("note\n=1+1\n=HYPERLINK(\"http://x\")\nhello\n")
    _, wb = run(p, kind="csv")
    c = wb.worksheets[0]["A2"]
    assert c.value == "=1+1" and c.data_type == "s"
    assert wb.worksheets[0]["A3"].data_type == "s"


def test_csv_leading_zeros_preserved(run, make_csv):
    p = make_csv("zip,qty\n01234,1\n00501,2\n")
    _, wb = run(p, kind="csv")
    assert wb.worksheets[0]["A2"].value == "01234"
    assert wb.worksheets[0]["B2"].value == 1


def _xlsx_bytes():
    wb = Workbook()
    ws = wb.active
    ws.append(["n"])
    ws.append(["1"])
    ws.append(["2"])
    b = BytesIO()
    wb.save(b)
    return b.getvalue()


def test_api_upload_download_regenerate(client):
    r = client.post("/api/jobs", files={"file": ("My Data.xlsx", _xlsx_bytes())})
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["download_name"] == "My Data_corrected.xlsx"
    d = client.get(body["download_url"])
    assert d.status_code == 200
    assert "My%20Data_corrected.xlsx" in d.headers["content-disposition"]
    assert load_workbook(BytesIO(d.content)).active["A2"].value == 1
    r2 = client.post(f"/api/jobs/{body['job_id']}/regenerate",
                     json={"overrides": {"Sheet": {"1": "Text"}}})
    assert r2.status_code == 200
    d2 = client.get(body["download_url"])
    assert load_workbook(BytesIO(d2.content)).active["A2"].value == "1"
    assert client.delete(f"/api/jobs/{body['job_id']}").status_code == 200
    assert client.get(body["download_url"]).status_code == 404


def test_api_rejects_bad_files(client):
    assert client.post("/api/jobs", files={"file": ("x.txt", b"hi")}).status_code == 415
    assert client.post("/api/jobs", files={"file": ("x.xlsx", b"not a zip")}).status_code == 400
    assert client.post("/api/jobs", files={"file": ("x.csv", b"")}).status_code == 400
    assert client.post("/api/jobs", files={"file": ("x.csv", b"a\x00b\x00")}).status_code == 400
    big = b"a,b\n" + b"1,2\n" * (1024 * 300)
    assert client.post("/api/jobs", files={"file": ("big.csv", big)}).status_code == 413
    # failed uploads leave no jobs behind
    assert client.app.state.store.jobs == {}


def test_api_unknown_job_and_bad_options(client):
    assert client.get("/api/jobs/nope/download").status_code == 404
    r = client.post("/api/jobs", data={"date_order": "zzz"}, files={"file": ("a.csv", b"a\n1\n")})
    assert r.status_code == 422
