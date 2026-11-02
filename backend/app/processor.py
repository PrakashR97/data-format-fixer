"""Column type detection and conversion for every worksheet (pure Python rules)."""
from __future__ import annotations

import os
import re
import time as _time
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .config import Settings
from .csv_reader import read_csv
from .detection import NULL_MARKERS, NumberParse, DateParse, parse_date, parse_number
from .errors import UserError

CONFIDENCE = 0.9
OVERRIDE_TYPES = {"Integer", "Decimal", "Date", "DateTime", "Text"}
DATE_ORDERS = {"auto", "dmy", "mdy"}
OUTPUT_FORMATS = {"mm/dd/yyyy", "dd/mm/yyyy", "yyyy-mm-dd"}
MAX_LOCATIONS = 25

# header words that suggest values are identifiers, not quantities
ID_WORDS = {
    "phone", "tel", "telephone", "mobile", "cell", "fax", "zip", "zipcode", "postal", "postcode",
    "pincode", "pin", "account", "acct", "id", "uuid", "guid", "ssn", "sku", "barcode", "isbn",
    "iban", "code", "ref", "reference", "serial", "passport", "licence", "license", "routing",
    "swift", "ein", "tin", "vat",
}


@dataclass
class Options:
    date_order: str = "auto"
    output_format: str = "yyyy-mm-dd"


@dataclass
class Item:
    row: int
    col: int
    raw: str
    kind: str  # null | int | dec | pnum | date | dbad | text
    num: Optional[NumberParse] = None
    date: Optional[DateParse] = None


@dataclass
class ColumnResult:
    index: int
    header: Optional[str]
    detected: str = "Text"
    effective: str = "Text"
    overridden: bool = False
    confidence: float = 1.0
    non_empty: int = 0
    converted: int = 0
    notes: list[str] = field(default_factory=list)
    samples: list[str] = field(default_factory=list)
    warning_count: int = 0

    def as_dict(self) -> dict:
        return {
            "index": self.index,
            "letter": get_column_letter(self.index),
            "header": self.header,
            "detected_type": self.detected,
            "effective_type": self.effective,
            "overridden": self.overridden,
            "confidence": round(self.confidence, 3),
            "non_empty": self.non_empty,
            "converted": self.converted,
            "warning_count": self.warning_count,
            "notes": self.notes,
            "samples": self.samples,
        }


class Deadline:
    def __init__(self, seconds: int):
        self.end = _time.monotonic() + seconds

    def check(self) -> None:
        if _time.monotonic() > self.end:
            raise UserError(
                "Processing took longer than the allowed time limit. Try a smaller file.", 413
            )


# --------------------------------------------------------------- helpers
def _tokens(header: str) -> set[str]:
    spaced = re.sub(r"([a-z])([A-Z])", r"\1 \2", header)
    return {t for t in re.split(r"[^A-Za-z0-9]+", spaced.lower()) if t}


def header_hint(header: Optional[str]) -> Optional[str]:
    if not header:
        return None
    hit = _tokens(header) & ID_WORDS
    return sorted(hit)[0] if hit else None


def classify(row: int, col: int, raw: str) -> Item:
    s = raw.strip()
    if not s or s.lower() in NULL_MARKERS:
        return Item(row, col, raw, "null")
    n = parse_number(s)
    if n.kind in ("int", "dec"):
        return Item(row, col, raw, n.kind, num=n)
    if n.kind == "protected":
        return Item(row, col, raw, "pnum", num=n)
    d = parse_date(s)
    if d.status in ("fixed", "numeric"):
        return Item(row, col, raw, "date", date=d)
    if d.status == "invalid":
        return Item(row, col, raw, "dbad", date=d)
    return Item(row, col, raw, "text")


def _short(v: str) -> str:
    v = v.strip()
    return v if len(v) <= 60 else v[:57] + "..."


def _native_kind(v: Any) -> str:
    if isinstance(v, bool):
        return "other"
    if isinstance(v, int):
        return "int"
    if isinstance(v, float):
        return "dec"
    if isinstance(v, datetime):
        return "datetime" if v.time() != time(0) else "date"
    if isinstance(v, date):
        return "date"
    return "other"


def _date_base(fmt: str) -> str:
    return fmt


class WarningBook:
    """Groups warnings by (sheet, column, reason) and keeps only a few locations."""

    def __init__(self) -> None:
        self.groups: dict[tuple, dict] = {}

    def add(self, sheet: str, col: ColumnResult, reason: str, row: int, raw: str) -> None:
        key = (sheet, col.index, reason)
        g = self.groups.get(key)
        if g is None:
            g = self.groups[key] = {
                "sheet": sheet,
                "column": get_column_letter(col.index),
                "header": col.header,
                "reason": reason,
                "ambiguous": reason.startswith("ambiguous"),
                "count": 0,
                "locations": [],
            }
        g["count"] += 1
        col.warning_count += 1
        if len(g["locations"]) < MAX_LOCATIONS:
            g["locations"].append(
                {"cell": f"{get_column_letter(col.index)}{row}", "value": _short(raw)}
            )

    def to_list(self) -> list[dict]:
        return list(self.groups.values())


# ---------------------------------------------------------- column logic
def _detect(
    items: list[Item], nat: Counter, header: Optional[str]
) -> tuple[str, float, list[str]]:
    k = Counter(i.kind for i in items)
    n_int = k["int"] + nat["int"]
    n_dec = k["dec"] + nat["dec"]
    n_pn = k["pnum"]
    n_date = k["date"] + k["dbad"] + nat["date"] + nat["datetime"]
    n_text = k["text"] + nat["other"]
    total = n_int + n_dec + n_pn + n_date + n_text
    notes: list[str] = []
    if total == 0:
        return "Text", 1.0, notes
    num_like = n_int + n_dec + n_pn
    if num_like / total >= CONFIDENCE:
        str_nums = [i for i in items if i.kind in ("int", "dec", "pnum")]
        if str_nums:
            if n_pn:
                notes.append(
                    "Contains identifier-like numbers (leading zeros, '+' prefix or more than "
                    "15 digits), so the whole column is kept as text."
                )
                return "Text", num_like / total, notes
            hint = header_hint(header)
            if hint:
                notes.append(
                    f"Header contains '{hint}', so the numbers are treated as identifiers "
                    "and kept as text."
                )
                return "Text", num_like / total, notes
            lengths = {len(i.raw.strip()) for i in str_nums}
            if (
                k["dec"] == 0
                and len(str_nums) >= 3
                and len(lengths) == 1
                and next(iter(lengths)) >= 8
                and not nat["int"]
                and not nat["dec"]
            ):
                notes.append(
                    "All values are digit strings of the same long length (looks like "
                    "identifiers), so they are kept as text."
                )
                return "Text", num_like / total, notes
        return ("Decimal" if n_dec else "Integer"), num_like / total, notes
    if n_date / total >= CONFIDENCE:
        has_time = nat["datetime"] > 0 or any(
            i.kind == "date" and i.date and i.date.tm for i in items
        )
        return ("DateTime" if has_time else "Date"), n_date / total, notes
    if n_text / total >= CONFIDENCE:
        return "Text", n_text / total, notes
    parts = []
    if num_like:
        parts.append(f"{round(100 * num_like / total)}% numbers")
    if n_date:
        parts.append(f"{round(100 * n_date / total)}% dates")
    if n_text:
        parts.append(f"{round(100 * n_text / total)}% text")
    notes.append("Column mixes types (" + ", ".join(parts) + "); nothing was converted.")
    return "Mixed", max(num_like, n_date, n_text) / total, notes


def _resolve_dates(
    items: list[Item], order: str, notes: list[str]
) -> dict[int, tuple[Optional[date], str]]:
    """Return {id(item): (date or None, reason-if-none)} for items of kind 'date'."""
    dated = [i for i in items if i.kind == "date"]
    out: dict[int, tuple[Optional[date], str]] = {}
    col_order = order if order in ("dmy", "mdy") else None
    conflict = False
    if order == "auto":
        only_dmy = sum(1 for i in dated if i.date.status == "numeric" and i.date.dmy and not i.date.mdy)
        only_mdy = sum(1 for i in dated if i.date.status == "numeric" and i.date.mdy and not i.date.dmy)
        if only_dmy and not only_mdy:
            col_order = "dmy"
            notes.append(
                f"Day-first order inferred from {only_dmy} unambiguous value(s) in this column."
            )
        elif only_mdy and not only_dmy:
            col_order = "mdy"
            notes.append(
                f"Month-first order inferred from {only_mdy} unambiguous value(s) in this column."
            )
        elif only_dmy and only_mdy:
            conflict = True
            notes.append("Column mixes day-first and month-first dates; ambiguous ones were kept.")
    for it in dated:
        dp = it.date
        if dp.status == "fixed":
            out[id(it)] = (dp.fixed, "")
        elif col_order:
            d = dp.dmy if col_order == "dmy" else dp.mdy
            label = "day-first" if col_order == "dmy" else "month-first"
            if d is None and order in ("dmy", "mdy"):
                # User-chosen order is a preference: a value that is valid in only
                # the other order is still unambiguous, so convert it.
                d = dp.mdy if col_order == "dmy" else dp.dmy
            out[id(it)] = (d, "" if d else f"not a valid {label} date")
        elif dp.dmy and not dp.mdy:
            out[id(it)] = (dp.dmy, "")
        elif dp.mdy and not dp.dmy:
            out[id(it)] = (dp.mdy, "")
        else:
            reason = (
                "ambiguous date (could be day-first or month-first) - choose a date order"
                if not conflict
                else "ambiguous date in a column with mixed date orders"
            )
            out[id(it)] = (None, reason)
    return out


def _process_column(
    ws, sheet: str, col: ColumnResult, items: list[Item], nat: Counter, override: Optional[str],
    options: Options, warnings: WarningBook, deadline: Deadline,
) -> None:
    detected, conf, notes = _detect(items, nat, col.header)
    col.detected, col.confidence, col.notes = detected, conf, notes
    forced = override in OVERRIDE_TYPES
    eff = override if forced else detected
    col.overridden = forced
    col.effective = eff
    if forced:
        col.notes.append(f"Type set manually to {eff}.")

    def write(it: Item, value: Any, fmt: str) -> None:
        cell = ws.cell(row=it.row, column=it.col)
        cell.value = value
        cell.number_format = fmt
        col.converted += 1

    def warn(it: Item, reason: str) -> None:
        warnings.add(sheet, col, reason, it.row, it.raw)
        if reason.startswith("ambiguous"):
            # Flag the kept value inside the output file too (value is untouched).
            from openpyxl.comments import Comment
            cell = ws.cell(row=it.row, column=it.col)
            if cell.comment is None:
                cell.comment = Comment(
                    "AMBIGUOUS DATE: could be day-first or month-first. "
                    "Kept as text. Choose a date order and regenerate.",
                    "Data Format Fixer",
                )

    if eff == "Mixed":
        for it in items:
            if it.kind in ("int", "dec", "pnum", "date", "dbad"):
                warn(it, "kept as text because the column mixes types")
        return
    if eff == "Text":
        return

    if eff in ("Integer", "Decimal"):
        label = "integer" if eff == "Integer" else "number"
        for it in items:
            if it.kind == "null":
                continue
            kind, n = it.kind, it.num
            if kind == "pnum" and forced and n and not n.hard:
                n = parse_number(it.raw, allow_protected=True)
                kind = n.kind
            if kind == "int" or (kind == "dec" and eff == "Decimal"):
                write(it, n.value, n.number_format)
            elif kind == "dec":
                warn(it, "decimal value in an Integer column, kept as text")
            elif kind in ("pnum", "protected"):
                warn(it, n.reason)
            elif kind in ("date", "dbad"):
                warn(it, f"looks like a date, not a {label}")
            else:
                warn(it, f"not a valid {label}, kept as is")
        return

    # Date / DateTime
    resolved = _resolve_dates(items, options.date_order, col.notes)
    converted_dates = [(it, resolved[id(it)][0]) for it in items if it.kind == "date" and resolved[id(it)][0]]
    has_time = any(it.date.tm for it, _ in converted_dates)
    micro = any(it.date.tm and it.date.tm.microsecond for it, _ in converted_dates)
    fmt = options.output_format
    if has_time or eff == "DateTime":
        fmt += " hh:mm:ss" + (".000" if micro else "")
    for it in items:
        if it.kind == "null":
            continue
        if it.kind == "date":
            d, reason = resolved[id(it)]
            if d is None:
                warn(it, reason)
            else:
                write(it, datetime.combine(d, it.date.tm or time(0)), fmt)
        elif it.kind == "dbad":
            warn(it, it.date.reason)
        elif it.kind in ("int", "dec", "pnum"):
            warn(it, "number in a date column, not treated as a date serial")
        else:
            warn(it, "not recognised as a date, kept as is")


# ---------------------------------------------------------------- sheets
def _is_header_value(v: Any) -> bool:
    if not isinstance(v, str) or not v.strip():
        return False
    return parse_number(v).kind == "none" and parse_date(v).status == "none"


def _process_sheet(
    ws, settings: Settings, options: Options, overrides: dict[str, str],
    warnings: WarningBook, deadline: Deadline, cells_seen: int,
) -> tuple[dict, int]:
    rows, cols = ws.max_row or 0, ws.max_column or 0
    if rows > settings.max_rows + 1:
        raise UserError(f"Sheet '{ws.title}' has more than {settings.max_rows:,} rows, which is the limit.")
    if cols > settings.max_columns:
        raise UserError(f"Sheet '{ws.title}' has more than {settings.max_columns} columns, which is the limit.")
    cells_seen += rows * cols
    if cells_seen > settings.max_cells:
        raise UserError(f"The workbook has more than {settings.max_cells:,} cells, which is the limit.", 413)
    if rows == 0 or (rows == 1 and cols == 1 and ws.cell(1, 1).value is None):
        return {"name": ws.title, "rows": 0, "columns": []}, cells_seen

    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=rows, max_col=cols)]
    first = next((i for i, r in enumerate(grid) if any(c.value not in (None, "") for c in r)), None)
    if first is None:
        return {"name": ws.title, "rows": 0, "columns": []}, cells_seen
    head_vals = [c.value for c in grid[first] if c.value not in (None, "")]
    has_header = all(_is_header_value(v) for v in head_vals)
    data_start = first + 1 if has_header else first

    results: list[ColumnResult] = []
    for c in range(cols):
        raw_header = grid[first][c].value if has_header else None
        header = str(raw_header).strip() if raw_header not in (None, "") else None
        col = ColumnResult(index=c + 1, header=header)
        items: list[Item] = []
        nat: Counter = Counter()
        for r in range(data_start, rows):
            cell = grid[r][c]
            if cell.data_type == "f" or type(cell).__name__ == "MergedCell":
                continue  # never touch formulas
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str):
                it = classify(r + 1, c + 1, v)
                if it.kind == "null" and not v.strip():
                    continue
                items.append(it)
            else:
                nat[_native_kind(v)] += 1
            if (r & 1023) == 0:
                deadline.check()
        col.non_empty = len(items) + sum(nat.values())
        col.samples = [_short(i.raw) for i in items if i.raw.strip()][:3]
        _process_column(ws, ws.title, col, items, nat, overrides.get(str(c + 1)), options, warnings, deadline)
        results.append(col)
        deadline.check()
    return {"name": ws.title, "rows": max(0, rows - data_start), "columns": [r.as_dict() for r in results]}, cells_seen


# --------------------------------------------------------------- entry
def validate_xlsx(path: Path, settings: Settings) -> list[str]:
    """Check the file really is an xlsx package; return notices about unsupported features."""
    if not zipfile.is_zipfile(path):
        raise UserError("This is not a valid .xlsx file (it is not an Excel package).")
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        names = {i.filename for i in infos}
        if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
            raise UserError("This file is not a valid Excel workbook (.xlsx).")
        if len(infos) > 20000 or sum(i.file_size for i in infos) > settings.max_uncompressed_mb * 1024 * 1024:
            raise UserError("The workbook is too large when unpacked.", 413)
    notices = []
    lost = []
    if any(n.startswith("xl/charts/") for n in names):
        lost.append("charts")
    if any(n.startswith("xl/pivotTables/") for n in names):
        lost.append("pivot tables")
    if any(n.startswith(("xl/slicers/", "xl/ctrlProps/", "xl/embeddings/")) for n in names):
        lost.append("slicers/form controls/embedded objects")
    if lost:
        notices.append(
            "This workbook contains " + ", ".join(lost) + ", which cannot be carried over to the "
            "corrected file. Your original file is untouched."
        )
    return notices


def process_file(
    source: Path, kind: str, filename: str, out_path: Path, settings: Settings,
    options: Options, overrides: dict[str, dict[str, str]],
) -> dict:
    """Read ``source``, convert, and write ``out_path``. Never modifies ``source``."""
    deadline = Deadline(settings.max_processing_seconds)
    notices: list[str] = []
    csv_info = None
    if kind == "csv":
        wb, csv_info, notices = read_csv(source, filename, settings)
    else:
        notices = validate_xlsx(source, settings)
        try:
            wb = load_workbook(source)
        except Exception as exc:  # corrupted, encrypted, etc.
            raise UserError(
                "The Excel file could not be read. It may be corrupted or password-protected."
            ) from None
    if len(wb.worksheets) > settings.max_sheets:
        raise UserError(f"The workbook has more than {settings.max_sheets} sheets, which is the limit.")

    warnings = WarningBook()
    sheets = []
    cells_seen = 0
    for ws in wb.worksheets:
        report, cells_seen = _process_sheet(
            ws, settings, options, overrides.get(ws.title, {}), warnings, deadline, cells_seen
        )
        sheets.append(report)

    tmp = out_path.with_suffix(".tmp")
    wb.save(tmp)
    os.replace(tmp, out_path)
    wb.close()
    warning_list = warnings.to_list()
    return {
        "sheets": sheets,
        "warnings": warning_list,
        "warning_total": sum(w["count"] for w in warning_list),
        "ambiguous_total": sum(w["count"] for w in warning_list if w.get("ambiguous")),
        "converted_total": sum(c["converted"] for s in sheets for c in s["columns"]),
        "notices": notices,
        "csv": csv_info,
    }
