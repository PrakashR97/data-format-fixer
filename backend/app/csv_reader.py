"""Read CSV files (common encodings/delimiters) into an openpyxl workbook of literal text."""
from __future__ import annotations

import csv
import io
import re
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from .config import Settings
from .errors import UserError

DELIMITERS = [",", ";", "\t", "|"]
DELIMITER_NAMES = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}
EXCEL_MAX_CELL_CHARS = 32767


def decode_bytes(raw: bytes) -> tuple[str, str]:
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8"), "UTF-8 (with BOM)"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            return raw.decode("utf-16"), "UTF-16"
        except UnicodeDecodeError:
            raise UserError("The CSV file looks like UTF-16 but could not be decoded.")
    if b"\x00" in raw:
        raise UserError("This does not look like a text CSV file (it contains binary data).")
    try:
        return raw.decode("utf-8"), "UTF-8"
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("cp1252"), "Windows-1252"
    except UnicodeDecodeError:
        return raw.decode("latin-1"), "Latin-1"


def detect_delimiter(text: str) -> str:
    sample = text[:100_000]
    lines = sample.splitlines()[:30]
    best, best_score = ",", 0
    for d in DELIMITERS:
        try:
            rows = [r for r in csv.reader(io.StringIO("\n".join(lines)), delimiter=d) if r]
        except csv.Error:
            continue
        if not rows:
            continue
        modal, freq = Counter(len(r) for r in rows).most_common(1)[0]
        score = freq * (modal - 1) if modal > 1 else 0
        if score > best_score:
            best, best_score = d, score
    return best


def _sheet_title(filename: str) -> str:
    stem = Path(filename).stem
    stem = re.sub(r"[\[\]:*?/\\]", "_", stem).strip("' ")[:31]
    return stem or "Sheet1"


def read_csv(path: Path, filename: str, settings: Settings) -> tuple[Workbook, dict, list[str]]:
    text, encoding = decode_bytes(path.read_bytes())
    delimiter = detect_delimiter(text)
    csv.field_size_limit(1_000_000)
    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_title(filename)
    removed = 0
    max_cols = 0
    try:
        for r, row in enumerate(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter), 1):
            if r > settings.max_rows + 1:
                raise UserError(f"The file has more than {settings.max_rows:,} rows, which is the limit.")
            if len(row) > settings.max_columns:
                raise UserError(f"The file has more than {settings.max_columns} columns, which is the limit.")
            max_cols = max(max_cols, len(row))
            for c, value in enumerate(row, 1):
                if value == "":
                    continue
                if len(value) > EXCEL_MAX_CELL_CHARS:
                    raise UserError(f"A cell in row {r} is longer than Excel's 32,767 character limit.")
                cleaned = ILLEGAL_CHARACTERS_RE.sub("", value)
                removed += len(value) - len(cleaned)
                cell = ws.cell(row=r, column=c, value=cleaned)
                cell.data_type = "s"  # always literal text, even if it starts with '='
    except csv.Error:
        raise UserError("The CSV file could not be parsed. It may be malformed.")
    notices = []
    if removed:
        notices.append(f"Removed {removed} invalid control character(s) that Excel cannot store.")
    info = {
        "encoding": encoding,
        "delimiter": DELIMITER_NAMES.get(delimiter, delimiter),
        "note": "CSV files are plain text: they cannot store Excel cell types or formatting, "
        "so types are detected from the text and the result is saved as an Excel (.xlsx) file.",
    }
    return wb, info, notices
