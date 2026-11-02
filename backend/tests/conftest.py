import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.config import Settings
from app.main import create_app
from app.processor import Options, process_file


@pytest.fixture
def settings():
    return Settings(max_upload_mb=1, max_rows=1000, max_cells=100000)


@pytest.fixture
def run(tmp_path, settings):
    """run(rows_or_workbook_path, kind, options, overrides) -> (report, output workbook)"""
    def _run(source, kind="xlsx", options=None, overrides=None, name="t"):
        out = tmp_path / "out.xlsx"
        report = process_file(source, kind, name + "." + kind, out, settings,
                              options or Options(), overrides or {})
        return report, load_workbook(out)
    return _run


@pytest.fixture
def make_xlsx(tmp_path):
    def _make(sheets: dict):
        wb = Workbook()
        wb.remove(wb.active)
        for title, rows in sheets.items():
            ws = wb.create_sheet(title)
            for r in rows:
                ws.append(r)
        p = tmp_path / "in.xlsx"
        wb.save(p)
        return p
    return _make


@pytest.fixture
def make_csv(tmp_path):
    def _make(text, encoding="utf-8"):
        p = tmp_path / "in.csv"
        p.write_bytes(text.encode(encoding))
        return p
    return _make


@pytest.fixture
def client(settings):
    with TestClient(create_app(settings)) as c:
        yield c
