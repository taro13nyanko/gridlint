"""Excel value model: numbers, text, booleans, errors, blanks, and coercion.

Excel's coercion rules are unusual enough that getting them wrong silently
changes results, so they live here alone and are unit-tested directly.
"""
from __future__ import annotations

import datetime as _dt
import math
from dataclasses import dataclass
from decimal import ROUND_CEILING, ROUND_FLOOR, ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

DIV0 = "#DIV/0!"
VALUE = "#VALUE!"
REF = "#REF!"
NAME = "#NAME?"
NUM = "#NUM!"
NA = "#N/A"


@dataclass(frozen=True)
class ExcelError(Exception):
    code: str

    def __str__(self) -> str:
        return self.code


class Blank:
    """Excel's empty cell. Coerces to 0 in arithmetic and "" in text."""
    _inst = None

    def __new__(cls):
        if cls._inst is None:
            cls._inst = super().__new__(cls)
        return cls._inst

    def __repr__(self) -> str:
        return "BLANK"

    def __bool__(self) -> bool:
        return False


BLANK = Blank()

Value = float | int | str | bool | ExcelError | Blank


def is_error(v: Any) -> bool:
    return isinstance(v, ExcelError)


#: Excel stores a date as the number of days since 1899-12-30. The offset is two
#: rather than one because Excel deliberately kept Lotus 1-2-3's belief that 1900
#: was a leap year, so serial 60 is a day that never existed.
_EPOCH = _dt.date(1899, 12, 30)
_LEAP_BUG_SERIAL = 60


def to_serial(v: _dt.date | _dt.datetime | _dt.time | _dt.timedelta) -> float:
    """Convert a Python date/time to the number Excel would store."""
    if isinstance(v, _dt.timedelta):
        return v.total_seconds() / 86400.0
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400.0
    if isinstance(v, _dt.datetime):
        days = (v.date() - _EPOCH).days
        frac = (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400.0
    else:
        days, frac = (v - _EPOCH).days, 0.0
    if days < _LEAP_BUG_SERIAL:
        days -= 1                      # dates before 1900-03-01 are one off, by design
    return days + frac


def from_serial(n: float) -> _dt.datetime:
    """Convert an Excel serial back to a datetime."""
    if n != n or n < 0:
        raise ExcelError(NUM)
    days = int(n)
    if days < _LEAP_BUG_SERIAL:
        days += 1
    frac = n - int(n)
    try:
        base = _EPOCH + _dt.timedelta(days=days)
    except OverflowError as e:
        raise ExcelError(NUM) from e
    return _dt.datetime(base.year, base.month, base.day) + _dt.timedelta(seconds=round(frac * 86400))


def to_number(v: Any) -> float:
    """Coerce to a number the way an arithmetic operator would."""
    if isinstance(v, ExcelError):
        raise v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        return to_serial(v)
    if isinstance(v, Blank):
        return 0.0
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return 0.0
        try:
            if s.endswith("%"):
                return float(s[:-1].replace(",", "")) / 100.0
            return float(s.replace(",", ""))
        except ValueError as e:
            raise ExcelError(VALUE) from e
    raise ExcelError(VALUE)


def general_format(x: float) -> str:
    """Excel's General number format, close enough for text concatenation."""
    if x != x or math.isinf(x):
        raise ExcelError(NUM)
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = repr(round(x, 15))
    if "e" in s or "E" in s:
        return f"{x:.10G}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def to_text(v: Any) -> str:
    if isinstance(v, ExcelError):
        raise v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, Blank):
        return ""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d" if not isinstance(v, _dt.datetime) or
                          (v.hour or v.minute or v.second) == 0 else "%Y-%m-%d %H:%M:%S")
    if isinstance(v, (int, float)):
        return general_format(float(v))
    return str(v)


def to_bool(v: Any) -> bool:
    if isinstance(v, ExcelError):
        raise v
    if isinstance(v, bool):
        return v
    if isinstance(v, Blank):
        return False
    if isinstance(v, (int, float)):
        return float(v) != 0.0
    if isinstance(v, str):
        s = v.strip().upper()
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        raise ExcelError(VALUE)
    raise ExcelError(VALUE)


# Excel orders types: number < text < FALSE < TRUE
def _rank(v: Any) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, (int, float)):
        return 0
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        return 0                                   # a date is a number to Excel
    return 1


def compare(a: Any, b: Any) -> int:
    """Return -1/0/1 using Excel's cross-type ordering."""
    if isinstance(a, ExcelError):
        raise a
    if isinstance(b, ExcelError):
        raise b
    if isinstance(a, Blank):
        a = 0.0 if isinstance(b, (int, float)) and not isinstance(b, bool) else ("" if isinstance(b, str) else 0.0)
    if isinstance(b, Blank):
        b = 0.0 if isinstance(a, (int, float)) and not isinstance(a, bool) else ("" if isinstance(a, str) else 0.0)
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 0:
        fa, fb = to_number(a), to_number(b)
        return 0 if fa == fb else (-1 if fa < fb else 1)
    if ra == 1:
        sa, sb = str(a).upper(), str(b).upper()   # Excel text comparison is case-insensitive
        return 0 if sa == sb else (-1 if sa < sb else 1)
    ba, bb = bool(a), bool(b)
    return 0 if ba == bb else (-1 if bb else 1)


def excel_round(x: float, digits: int = 0, mode: str = "half_up") -> float:
    """Excel rounds half AWAY FROM ZERO; Python's round() is banker's rounding.

    The rounding is done in decimal, on the shortest string that round-trips to
    the same float. That is the number the person typing into the cell believes
    they have, and it is what Excel rounds: ROUND(2.675,2) is 2.68 even though
    2.675 is stored as 2.67499999999999982236431605997495353221893310546875.

    An earlier version nudged the scaled value by a relative epsilon instead.
    That worked for 2.675 but pushed 3806241.4967 over the boundary to
    3806242, which the clean-corpus benchmark caught.
    """
    if x != x or math.isinf(x):
        raise ExcelError(NUM)
    if mode not in ("half_up", "up", "down"):
        raise ValueError(f"unknown rounding mode {mode!r}")
    try:
        d = Decimal(repr(float(x)))
        q = Decimal(1).scaleb(-digits)
        if mode == "half_up":
            out = d.quantize(q, rounding=ROUND_HALF_UP)
        else:
            # ROUNDUP and ROUNDDOWN move away from / toward zero, not toward infinity.
            sign = Decimal(-1) if d < 0 else Decimal(1)
            out = (abs(d).quantize(q, rounding=ROUND_CEILING if mode == "up" else ROUND_FLOOR)) * sign
        return float(out)
    except (InvalidOperation, ValueError, OverflowError) as e:
        raise ExcelError(NUM) from e
