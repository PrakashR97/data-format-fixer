"""Rule-based parsers for numbers and dates stored as text. No AI involved."""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, time
from typing import Optional

NULL_MARKERS = {"n/a", "na", "null", "none", "nan", "-", "--", "#n/a", "nil", "(null)", "(none)"}

# Excel keeps 15 significant digits; more than that would be silently rounded.
MAX_SIG_DIGITS = 15
MIN_YEAR = 1900  # Excel cannot represent earlier dates

# ---------------------------------------------------------------- numbers
_PLAIN = re.compile(r"^(?P<sign>[+-])?(?P<int>\d+)(?:\.(?P<frac>\d+))?$")
_LEADING_DOT = re.compile(r"^(?P<sign>[+-])?\.(?P<frac>\d+)$")
_THOUSANDS = re.compile(r"^(?P<sign>[+-])?(?P<int>\d{1,3}(?:,\d{3})+)(?:\.(?P<frac>\d+))?$")


@dataclass
class NumberParse:
    kind: str  # 'int' | 'dec' | 'protected' | 'none'
    value: Optional[float | int] = None
    number_format: str = "General"
    reason: str = ""
    hard: bool = False  # hard = cannot be stored in Excel without losing precision


def parse_number(text: str, allow_protected: bool = False) -> NumberParse:
    """Parse numeric text.

    Values that look like identifiers (leading zeros, '+' prefix) are 'protected'
    unless ``allow_protected`` is set (explicit user override). Values needing more
    than 15 significant digits are always protected: Excel would round them.
    """
    s = text.strip()
    m = _THOUSANDS.match(s) or _PLAIN.match(s) or _LEADING_DOT.match(s)
    if not m:
        return NumberParse("none")
    groups = m.groupdict()
    sign = groups.get("sign") or ""
    int_digits = (groups.get("int") or "").replace(",", "")
    frac = groups.get("frac") or ""
    has_commas = "," in s

    int_sig = int_digits.lstrip("0")
    if int_sig:
        sig = len(int_sig) + len(frac.rstrip("0"))
    else:
        sig = len(frac.strip("0"))
    if sig > MAX_SIG_DIGITS:
        return NumberParse(
            "protected",
            reason="more than 15 significant digits - Excel would round it, so it is kept as text",
            hard=True,
        )
    if not allow_protected:
        if len(int_digits) > 1 and int_digits.startswith("0"):
            return NumberParse("protected", reason="has leading zeros (identifier-like), kept as text")
        if sign == "+":
            return NumberParse("protected", reason="starts with '+' (phone-number-like), kept as text")

    if frac:
        fmt = ("#,##0." if has_commas else "0.") + "0" * len(frac)
        return NumberParse("dec", float(f"{sign}{int_digits or '0'}.{frac}"), fmt)
    return NumberParse("int", int(f"{sign}{int_digits}"), "#,##0" if has_commas else "General")


# ------------------------------------------------------------------ dates
_MONTHS: dict[str, int] = {}
for _i in range(1, 13):
    _MONTHS[calendar.month_name[_i].lower()] = _i
    _MONTHS[calendar.month_abbr[_i].lower()] = _i
_MONTHS["sept"] = 9

_TIME = r"(?P<h>\d{1,2}):(?P<mi>\d{2})(?::(?P<s>\d{2})(?:[.,](?P<f>\d{1,6}))?)?"
_TIME_TAIL = re.compile(rf"^(?P<date>.+?)(?:T|\s+){_TIME}\s*(?P<ampm>[AaPp][Mm])?$")
_TZ_TAIL = re.compile(
    rf"^(?P<date>.+?)(?:T|\s+){_TIME}\s*(?:Z|UTC|GMT|[+-]\d{{2}}:?\d{{2}})$", re.IGNORECASE
)
_ISO = re.compile(r"^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$")
_NUMERIC = re.compile(r"^(\d{1,2})([-/.])(\d{1,2})\2(\d{4})$")
_DAY_MONTH_NAME = re.compile(r"^(\d{1,2})(?:st|nd|rd|th)?[ -]([A-Za-z]{3,9})\.?,?[ -](\d{4})$")
_MONTH_NAME_DAY = re.compile(r"^([A-Za-z]{3,9})\.?[ -](\d{1,2})(?:st|nd|rd|th)?,?[ -](\d{4})$")


@dataclass
class DateParse:
    status: str  # 'none' | 'fixed' | 'numeric' | 'invalid'
    fixed: Optional[date] = None
    dmy: Optional[date] = None  # reading as day/month/year
    mdy: Optional[date] = None  # reading as month/day/year
    tm: Optional[time] = None
    reason: str = ""


def _safe(y: int, m: int, d: int) -> Optional[date]:
    if y < MIN_YEAR:
        return None
    try:
        return date(y, m, d)
    except ValueError:
        return None


def _parse_date_part(part: str) -> DateParse:
    part = part.strip()
    m = _ISO.match(part)
    if m:
        d = _safe(int(m[1]), int(m[3]), int(m[4]))
        if d:
            return DateParse("fixed", fixed=d)
        return DateParse("invalid", reason="not a real calendar date (or before year 1900)")
    m = _NUMERIC.match(part)
    if m:
        a, b, y = int(m[1]), int(m[3]), int(m[4])
        dmy, mdy = _safe(y, b, a), _safe(y, a, b)
        if dmy is None and mdy is None:
            return DateParse("invalid", reason="not a real calendar date (or before year 1900)")
        if dmy == mdy:
            return DateParse("fixed", fixed=dmy)
        return DateParse("numeric", dmy=dmy, mdy=mdy)
    for rx, order in ((_DAY_MONTH_NAME, "dmy"), (_MONTH_NAME_DAY, "mdy")):
        m = rx.match(part)
        if m:
            if order == "dmy":
                day, name, year = int(m[1]), m[2], int(m[3])
            else:
                name, day, year = m[1], int(m[2]), int(m[3])
            month = _MONTHS.get(name.lower())
            if month is None:
                return DateParse("none")
            d = _safe(year, month, day)
            if d:
                return DateParse("fixed", fixed=d)
            return DateParse("invalid", reason="not a real calendar date (or before year 1900)")
    return DateParse("none")


def parse_date(text: str) -> DateParse:
    s = text.strip()
    tz = _TZ_TAIL.match(s)
    if tz:
        if _parse_date_part(tz["date"]).status in ("fixed", "numeric"):
            return DateParse(
                "invalid", reason="has a time-zone offset, which Excel cells cannot store"
            )
        return DateParse("none")
    tm: Optional[time] = None
    date_part = s
    m = _TIME_TAIL.match(s)
    if m:
        date_part = m["date"]
        h, mi = int(m["h"]), int(m["mi"])
        sec = int(m["s"]) if m["s"] else 0
        micro = int((m["f"] or "0").ljust(6, "0")) if m["f"] else 0
        ampm = (m["ampm"] or "").lower()
        if ampm:
            if not 1 <= h <= 12:
                return DateParse("invalid", reason="invalid time of day")
            h = h % 12 + (12 if ampm == "pm" else 0)
        if h > 23 or mi > 59 or sec > 59:
            result = _parse_date_part(date_part)
            if result.status in ("fixed", "numeric"):
                return DateParse("invalid", reason="invalid time of day")
            return DateParse("none")
        tm = time(h, mi, sec, micro)
    result = _parse_date_part(date_part)
    result.tm = tm
    return result
