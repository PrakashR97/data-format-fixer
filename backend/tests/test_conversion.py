import datetime as dt
from io import BytesIO

from app.processor import Options


def col(wb, sheet="S", letter="A", start=2):
    ws = wb[sheet]
    return [ws[f"{letter}{r}"].value for r in range(start, ws.max_row + 1)]


def test_numeric_strings_become_numbers(run, make_xlsx):
    p = make_xlsx({"S": [["n", "d"], ["1", "1.50"], ["2", "2.25"], ["1,234", "-3.1"], [None, None], ["5", "0.5"]]})
    report, wb = run(p)
    assert col(wb) == [1, 2, 1234, None, 5]
    assert all(isinstance(v, (int, type(None))) for v in col(wb))
    assert col(wb, letter="B") == [1.5, 2.25, -3.1, None, 0.5]
    assert wb["S"]["B2"].number_format == "0.00"  # decimals shown as written
    types = {c["header"]: c["effective_type"] for c in report["sheets"][0]["columns"]}
    assert types == {"n": "Integer", "d": "Decimal"}


def test_blank_cells_and_order_preserved(run, make_xlsx):
    p = make_xlsx({"S": [["n"], ["3"], [None], ["1"], ["2"]]})
    _, wb = run(p)
    assert col(wb) == [3, None, 1, 2]


def test_leading_zeros_stay_text(run, make_xlsx):
    p = make_xlsx({"S": [["val"], ["007"], ["0123"], ["45"]]})
    report, wb = run(p)
    assert col(wb) == ["007", "0123", "45"]
    assert report["sheets"][0]["columns"][0]["effective_type"] == "Text"


def test_header_hint_keeps_phone_and_zip_text(run, make_xlsx):
    p = make_xlsx({"S": [["Phone", "Zip"], ["5551234567", "90210"], ["5559876543", "10001"]]})
    _, wb = run(p)
    assert col(wb) == ["5551234567", "5559876543"]
    assert col(wb, letter="B") == ["90210", "10001"]


def test_long_identifiers_stay_text(run, make_xlsx):
    p = make_xlsx({"S": [["amount"], ["1"], ["12345678901234567890"]]})
    report, wb = run(p)
    assert "12345678901234567890" in col(wb)
    p2 = make_xlsx({"S": [["amount"]] + [["1"]] * 5 + [["1234567890123456"]]})
    report, wb = run(p2)
    assert col(wb)[-1] == "1234567890123456"


def test_long_value_not_converted_even_when_forced(run, make_xlsx):
    p = make_xlsx({"S": [["x"], ["1"], ["12345678901234567890"]]})
    report, wb = run(p, overrides={"S": {"1": "Integer"}})
    assert col(wb) == [1, "12345678901234567890"]
    assert report["warning_total"] == 1


def test_forced_integer_converts_leading_zeros(run, make_xlsx):
    p = make_xlsx({"S": [["val"], ["007"], ["45"]]})
    _, wb = run(p, overrides={"S": {"1": "Integer"}})
    assert col(wb) == [7, 45]


def test_ambiguous_dates_not_guessed(run, make_xlsx):
    p = make_xlsx({"S": [["d"], ["01/02/2026"], ["03/04/2026"]]})
    report, wb = run(p)
    assert col(wb) == ["01/02/2026", "03/04/2026"]
    assert report["warning_total"] == 2
    assert report["warnings"][0]["locations"][0]["cell"] == "A2"


def test_date_order_choice(run, make_xlsx):
    p = make_xlsx({"S": [["d"], ["01/02/2026"]]})
    _, wb = run(p, options=Options(date_order="dmy"))
    assert col(wb) == [dt.datetime(2026, 2, 1)]
    _, wb = run(p, options=Options(date_order="mdy", output_format="dd/mm/yyyy"))
    assert col(wb) == [dt.datetime(2026, 1, 2)]
    assert wb["S"]["A2"].number_format == "dd/mm/yyyy"


def test_auto_infers_order_from_unambiguous_value(run, make_xlsx):
    p = make_xlsx({"S": [["d"], ["13/02/2026"], ["01/02/2026"]]})
    report, wb = run(p)
    assert col(wb) == [dt.datetime(2026, 2, 13), dt.datetime(2026, 2, 1)]
    assert any("inferred" in n for n in report["sheets"][0]["columns"][0]["notes"])


def test_iso_dates_and_datetimes_keep_time(run, make_xlsx):
    p = make_xlsx({"S": [["d"], ["2026-03-05 14:30:15"], ["2026-03-06"], ["5 Mar 2026 2:05 PM"]]})
    report, wb = run(p)
    assert col(wb) == [dt.datetime(2026, 3, 5, 14, 30, 15), dt.datetime(2026, 3, 6), dt.datetime(2026, 3, 5, 14, 5)]
    assert report["sheets"][0]["columns"][0]["effective_type"] == "DateTime"
    assert "hh:mm:ss" in wb["S"]["A2"].number_format


def test_invalid_dates_preserved_and_reported(run, make_xlsx):
    p = make_xlsx({"S": [["d"]] + [["2026-01-%02d" % i] for i in range(1, 11)] + [["2026-02-30"]]})
    report, wb = run(p)
    assert col(wb)[-1] == "2026-02-30"
    w = report["warnings"][0]
    assert w["locations"][0]["cell"] == "A12"


def test_ordinary_numbers_never_become_dates(run, make_xlsx):
    p = make_xlsx({"S": [["value"], ["45000"], ["45001"]]})
    _, wb = run(p)
    assert col(wb) == [45000, 45001]
    assert wb["S"]["A2"].number_format == "General"


def test_mixed_column_untouched(run, make_xlsx):
    p = make_xlsx({"S": [["m"], ["1"], ["abc"], ["2"], ["def"], ["2026-01-01"]]})
    report, wb = run(p)
    assert col(wb) == ["1", "abc", "2", "def", "2026-01-01"]
    assert report["sheets"][0]["columns"][0]["effective_type"] == "Mixed"
    assert report["warning_total"] > 0


def test_mostly_numeric_column_converts_and_warns(run, make_xlsx):
    rows = [["n"]] + [[str(i)] for i in range(1, 20)] + [["oops"]]
    report, wb = run(make_xlsx({"S": rows}))
    assert col(wb)[:19] == list(range(1, 20))
    assert col(wb)[-1] == "oops"
    assert report["warnings"][0]["locations"][0] == {"cell": "A21", "value": "oops"}


def test_formulas_untouched(run, make_xlsx):
    p = make_xlsx({"S": [["a", "b"], ["1", "=A2*2"], ["2", "=SUM(A2:A3)"]]})
    _, wb = run(p)
    assert wb["S"]["B2"].value == "=A2*2"
    assert wb["S"]["B3"].value == "=SUM(A2:A3)"
    assert wb["S"]["A2"].value == 1


def test_multiple_sheets_names_and_order(run, make_xlsx):
    p = make_xlsx({"Zeta": [["n"], ["1"]], "Alpha": [["d"], ["2026-01-02"]], "Empty": []})
    report, wb = run(p)
    assert wb.sheetnames == ["Zeta", "Alpha", "Empty"]
    assert wb["Zeta"]["A2"].value == 1
    assert wb["Alpha"]["A2"].value == dt.datetime(2026, 1, 2)
    assert [s["name"] for s in report["sheets"]] == ["Zeta", "Alpha", "Empty"]


def test_existing_native_cells_and_rows_not_removed(run, make_xlsx):
    p = make_xlsx({"S": [["n"], [5], ["5"], ["5"]]})
    _, wb = run(p)
    assert col(wb) == [5, 5, 5]  # duplicates kept


def test_override_to_text_keeps_strings(run, make_xlsx):
    p = make_xlsx({"S": [["n"], ["1"], ["2"]]})
    _, wb = run(p, overrides={"S": {"1": "Text"}})
    assert col(wb) == ["1", "2"]


def test_original_file_not_modified(run, make_xlsx):
    p = make_xlsx({"S": [["n"], ["1"]]})
    before = p.read_bytes()
    run(p)
    assert p.read_bytes() == before


def test_chosen_date_order_resolves_ambiguous_but_keeps_invalid(tmp_path):
    import datetime as dt
    import openpyxl
    from app import processor
    from app.config import Settings

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Joining_Date"])
    for v in ["2026-01-15", "01/02/2026", "13/02/2026", "02/14/2026", "31/02/2026"]:
        ws.append([v])
    src = tmp_path / "a.xlsx"
    wb.save(src)

    def run(order):
        out = tmp_path / f"o_{order}.xlsx"
        processor.process_file(src, "xlsx", "a.xlsx", out, Settings(),
                               processor.Options(date_order=order), {})
        return [c.value for c in openpyxl.load_workbook(out).active["A"]][1:]

    auto = run("auto")
    assert auto[1] == "01/02/2026"  # never guessed in Auto
    mdy = run("mdy")
    assert mdy[1] == dt.datetime(2026, 1, 2)
    assert mdy[2] == dt.datetime(2026, 2, 13)  # only valid day-first, still converted
    assert mdy[3] == dt.datetime(2026, 2, 14)
    assert mdy[4] == "31/02/2026"  # invalid date stays text
    assert run("dmy")[1] == dt.datetime(2026, 2, 1)


def test_ambiguous_date_is_flagged(tmp_path):
    import openpyxl
    from app import processor
    from app.config import Settings

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["D"])
    for v in ["01/02/2026", "2026-03-04", "2026-05-06"]:
        ws.append([v])
    src = tmp_path / "a.xlsx"
    wb.save(src)
    out = tmp_path / "o.xlsx"
    r = processor.process_file(src, "xlsx", "a.xlsx", out, Settings(), processor.Options(), {})
    assert r["ambiguous_total"] == 1
    assert any(w["ambiguous"] for w in r["warnings"])
    cell = openpyxl.load_workbook(out).active["A2"]
    assert cell.value == "01/02/2026" and "AMBIGUOUS" in cell.comment.text
