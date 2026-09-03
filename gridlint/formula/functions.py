"""Worksheet functions.

Every function takes already-evaluated arguments. A range argument arrives as
a `Matrix` (list of rows of values) so aggregate functions can apply Excel's
"ignore text and blanks inside ranges, but coerce direct arguments" rule --
the single most common source of wrong reimplementations.
"""
from __future__ import annotations

import math
import re
from typing import Any, Callable

from .values import (BLANK, DIV0, NA, NUM, VALUE, Blank, ExcelError, compare,
                     excel_round, general_format, to_bool, to_number, to_text)


class Matrix(list):
    """A 2-D block of values produced by a range reference."""

    def flat(self):
        for row in self:
            for v in row:
                yield v


def _numbers_in(args, *, coerce_scalars: bool = True) -> list[float]:
    """Excel: inside a range, ignore text/blank/bool; a direct scalar is coerced."""
    out: list[float] = []
    for a in args:
        if isinstance(a, Matrix):
            for v in a.flat():
                if isinstance(v, ExcelError):
                    raise v
                if isinstance(v, bool) or isinstance(v, Blank) or isinstance(v, str):
                    continue
                out.append(float(v))
        elif isinstance(a, ExcelError):
            raise a
        elif coerce_scalars:
            if isinstance(a, Blank):
                continue
            out.append(to_number(a))
    return out


_CRIT_RE = re.compile(r"^(<=|>=|<>|=|<|>)?(.*)$", re.S)


def _criteria_pred(criteria: Any) -> Callable[[Any], bool]:
    """Build a predicate from a COUNTIF/SUMIF criteria value."""
    if isinstance(criteria, ExcelError):
        raise criteria
    if isinstance(criteria, (int, float)) and not isinstance(criteria, bool):
        target = float(criteria)
        return lambda v: (not isinstance(v, (str, Blank, bool))) and float(v) == target
    text = to_text(criteria)
    m = _CRIT_RE.match(text)
    op, rest = m.group(1) or "=", m.group(2).strip()
    try:
        rhs: Any = float(rest)
    except ValueError:
        rhs = rest
    wildcard = isinstance(rhs, str) and ("*" in rhs or "?" in rhs)
    if wildcard and op in ("=", "<>"):
        pat = re.compile("^" + re.escape(rhs).replace(r"\*", ".*").replace(r"\?", ".") + "$", re.I)
        return lambda v: bool(pat.match(to_text(v))) if op == "=" else not bool(pat.match(to_text(v)))
    if isinstance(rhs, str) and rhs == "" and op == "=":
        return lambda v: isinstance(v, Blank)

    def pred(v: Any) -> bool:
        if isinstance(v, ExcelError):
            return False
        if isinstance(v, Blank) and not isinstance(rhs, str):
            return False
        try:
            c = compare(v, rhs)
        except ExcelError:
            return False
        return {"=": c == 0, "<>": c != 0, "<": c < 0, "<=": c <= 0, ">": c > 0, ">=": c >= 0}[op]

    return pred


def _as_matrix(a: Any) -> Matrix:
    return a if isinstance(a, Matrix) else Matrix([[a]])


def f_sum(*args):
    return math.fsum(_numbers_in(args))


def f_average(*args):
    nums = _numbers_in(args)
    if not nums:
        raise ExcelError(DIV0)
    return math.fsum(nums) / len(nums)


def f_min(*args):
    nums = _numbers_in(args)
    return min(nums) if nums else 0.0


def f_max(*args):
    nums = _numbers_in(args)
    return max(nums) if nums else 0.0


def f_count(*args):
    n = 0
    for a in args:
        for v in (_as_matrix(a).flat() if isinstance(a, Matrix) else [a]):
            if isinstance(v, ExcelError) or isinstance(v, Blank) or isinstance(v, bool):
                continue
            if isinstance(v, (int, float)):
                n += 1
            elif isinstance(v, str) and not isinstance(a, Matrix):
                try:
                    to_number(v)
                    n += 1
                except ExcelError:
                    pass
    return float(n)


def f_counta(*args):
    n = 0
    for a in args:
        for v in (_as_matrix(a).flat() if isinstance(a, Matrix) else [a]):
            if not isinstance(v, Blank):
                n += 1
    return float(n)


def f_countblank(a):
    return float(sum(1 for v in _as_matrix(a).flat() if isinstance(v, Blank)))


def f_countif(rng, criteria):
    pred = _criteria_pred(criteria)
    return float(sum(1 for v in _as_matrix(rng).flat() if pred(v)))


def f_sumif(rng, criteria, sum_range=None):
    pred = _criteria_pred(criteria)
    src = _as_matrix(rng)
    tgt = _as_matrix(sum_range) if sum_range is not None else src
    total = 0.0
    flat_src = list(src.flat())
    flat_tgt = list(tgt.flat())
    for i, v in enumerate(flat_src):
        if not pred(v):
            continue
        if i < len(flat_tgt):
            t = flat_tgt[i]
            if isinstance(t, ExcelError):
                raise t
            if isinstance(t, (int, float)) and not isinstance(t, bool):
                total += float(t)
    return total


def f_averageif(rng, criteria, avg_range=None):
    pred = _criteria_pred(criteria)
    src = list(_as_matrix(rng).flat())
    tgt = list(_as_matrix(avg_range).flat()) if avg_range is not None else src
    vals = [float(tgt[i]) for i, v in enumerate(src)
            if pred(v) and i < len(tgt) and isinstance(tgt[i], (int, float)) and not isinstance(tgt[i], bool)]
    if not vals:
        raise ExcelError(DIV0)
    return math.fsum(vals) / len(vals)


def f_if(cond, then_v=True, else_v=False):
    return then_v if to_bool(cond) else else_v


def f_iferror(v, alt):
    return alt if isinstance(v, ExcelError) else v


def f_ifna(v, alt):
    return alt if isinstance(v, ExcelError) and v.code == NA else v


def f_iserror(v):
    return isinstance(v, ExcelError)


def f_isblank(v):
    if isinstance(v, Matrix):
        flat = list(v.flat())
        return len(flat) == 1 and isinstance(flat[0], Blank)
    return isinstance(v, Blank)


def f_isnumber(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def f_istext(v):
    return isinstance(v, str)


def f_abs(v):
    return abs(to_number(v))


def f_round(v, digits=0.0):
    return excel_round(to_number(v), int(to_number(digits)))


def f_roundup(v, digits=0.0):
    return excel_round(to_number(v), int(to_number(digits)), mode="up")


def f_rounddown(v, digits=0.0):
    return excel_round(to_number(v), int(to_number(digits)), mode="down")


def f_int(v):
    return float(math.floor(to_number(v)))


def f_sqrt(v):
    x = to_number(v)
    if x < 0:
        raise ExcelError(NUM)
    return math.sqrt(x)


def f_power(a, b):
    try:
        return float(to_number(a) ** to_number(b))
    except (OverflowError, ValueError, ZeroDivisionError) as e:
        raise ExcelError(NUM) from e


def f_product(*args):
    nums = _numbers_in(args)
    out = 1.0
    for n in nums:
        out *= n
    return out


def f_and(*args):
    vals = [v for a in args for v in (a.flat() if isinstance(a, Matrix) else [a])]
    return all(to_bool(v) for v in vals if not isinstance(v, Blank))


def f_or(*args):
    vals = [v for a in args for v in (a.flat() if isinstance(a, Matrix) else [a])]
    return any(to_bool(v) for v in vals if not isinstance(v, Blank))


def f_not(v):
    return not to_bool(v)


def f_concatenate(*args):
    return "".join(to_text(v) for a in args for v in (a.flat() if isinstance(a, Matrix) else [a]))


def f_len(v):
    return float(len(to_text(v)))


def f_left(v, n=1.0):
    return to_text(v)[: int(to_number(n))]


def f_right(v, n=1.0):
    n = int(to_number(n))
    return to_text(v)[-n:] if n else ""


def f_mid(v, start, n):
    s, i, ln = to_text(v), int(to_number(start)), int(to_number(n))
    if i < 1:
        raise ExcelError(VALUE)
    return s[i - 1: i - 1 + ln]


def f_trim(v):
    return " ".join(to_text(v).split())


def f_upper(v):
    return to_text(v).upper()


def f_lower(v):
    return to_text(v).lower()


def f_value(v):
    return to_number(v)


def f_text(v, fmt):
    """Minimal TEXT(): supports 0/#/. patterns and percent."""
    x, f = to_number(v), to_text(fmt)
    if "%" in f:
        decimals = len(f.split(".")[1].replace("%", "")) if "." in f else 0
        return f"{x * 100:.{decimals}f}%"
    decimals = len(f.split(".")[1]) if "." in f else 0
    comma = "," if "," in f.split(".")[0] else ""
    return f"{x:{comma}.{decimals}f}"


def _lookup_column(matrix: Matrix, index: int):
    if index < 1 or (matrix and index > len(matrix[0])):
        raise ExcelError(VALUE)
    return index - 1


def f_vlookup(needle, table, col_index, approximate=True):
    m = _as_matrix(table)
    ci = _lookup_column(m, int(to_number(col_index)))
    approx = to_bool(approximate) if not isinstance(approximate, bool) else approximate
    if not approx:
        for row in m:
            if compare(row[0], needle) == 0:
                return row[ci]
        raise ExcelError(NA)
    best = None
    for row in m:
        try:
            if compare(row[0], needle) <= 0:
                best = row
            else:
                break
        except ExcelError:
            continue
    if best is None:
        raise ExcelError(NA)
    return best[ci]


def f_index(array, row_num, col_num=1.0):
    m = _as_matrix(array)
    r, c = int(to_number(row_num)), int(to_number(col_num))
    if len(m) == 1 and c == 1 and r > 1:      # single-row array indexed linearly
        m0 = m[0]
        if r <= len(m0):
            return m0[r - 1]
    if r < 1 or r > len(m):
        raise ExcelError(REF_CODE)
    row = m[r - 1]
    if c < 1 or c > len(row):
        raise ExcelError(REF_CODE)
    return row[c - 1]


REF_CODE = "#REF!"


def f_match(needle, array, match_type=1.0):
    m = _as_matrix(array)
    flat = list(m.flat())
    mt = int(to_number(match_type))
    if mt == 0:
        for i, v in enumerate(flat):
            try:
                if compare(v, needle) == 0:
                    return float(i + 1)
            except ExcelError:
                continue
        raise ExcelError(NA)
    best = None
    for i, v in enumerate(flat):
        try:
            c = compare(v, needle)
        except ExcelError:
            continue
        if (mt > 0 and c <= 0) or (mt < 0 and c >= 0):
            best = i + 1
    if best is None:
        raise ExcelError(NA)
    return float(best)


FUNCTIONS: dict[str, Callable[..., Any]] = {
    "SUM": f_sum, "AVERAGE": f_average, "MIN": f_min, "MAX": f_max,
    "COUNT": f_count, "COUNTA": f_counta, "COUNTBLANK": f_countblank,
    "COUNTIF": f_countif, "SUMIF": f_sumif, "AVERAGEIF": f_averageif,
    "IF": f_if, "IFERROR": f_iferror, "IFNA": f_ifna,
    "ISERROR": f_iserror, "ISBLANK": f_isblank, "ISNUMBER": f_isnumber, "ISTEXT": f_istext,
    "ABS": f_abs, "ROUND": f_round, "ROUNDUP": f_roundup, "ROUNDDOWN": f_rounddown,
    "INT": f_int, "SQRT": f_sqrt, "POWER": f_power, "PRODUCT": f_product,
    "AND": f_and, "OR": f_or, "NOT": f_not,
    "CONCATENATE": f_concatenate, "CONCAT": f_concatenate,
    "LEN": f_len, "LEFT": f_left, "RIGHT": f_right, "MID": f_mid, "TRIM": f_trim,
    "UPPER": f_upper, "LOWER": f_lower, "VALUE": f_value, "TEXT": f_text,
    "VLOOKUP": f_vlookup, "INDEX": f_index, "MATCH": f_match,
}

#: Functions whose result depends on more than the workbook (recalculated every open).
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"}
