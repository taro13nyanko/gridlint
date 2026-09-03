"""Worksheet functions.

Every function takes already-evaluated arguments. A range argument arrives as
a `Matrix` (list of rows of values) so aggregate functions can apply Excel's
"ignore text and blanks inside ranges, but coerce direct arguments" rule --
the single most common source of wrong reimplementations.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
from typing import Any, Callable

from .values import (BLANK, DIV0, NA, NUM, VALUE, Blank, ExcelError, compare,
                     excel_round, from_serial, general_format, to_bool, to_number,
                     to_serial, to_text)


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


# ---------------------------------------------------------------------------
# Multi-criteria aggregates. These are what real operating models are built
# from, so a checker that cannot evaluate them cannot price anything downstream
# of them.
# ---------------------------------------------------------------------------

def _criteria_mask(pairs) -> list[bool]:
    """AND together (range, criteria) pairs, cell by cell."""
    masks: list[list[bool]] = []
    length = 0
    for rng, criteria in pairs:
        flat = list(_as_matrix(rng).flat())
        length = max(length, len(flat))
        pred = _criteria_pred(criteria)
        masks.append([pred(v) for v in flat])
    return [all(m[i] if i < len(m) else False for m in masks) for i in range(length)]


def _pairs(args) -> list[tuple[Any, Any]]:
    return [(args[i], args[i + 1]) for i in range(0, len(args) - 1, 2)]


def f_sumifs(sum_range, *args):
    mask = _criteria_mask(_pairs(args))
    flat = list(_as_matrix(sum_range).flat())
    return math.fsum(float(flat[i]) for i, keep in enumerate(mask)
                     if keep and i < len(flat)
                     and isinstance(flat[i], (int, float)) and not isinstance(flat[i], bool))


def f_countifs(*args):
    return float(sum(_criteria_mask(_pairs(args))))


def f_averageifs(avg_range, *args):
    mask = _criteria_mask(_pairs(args))
    flat = list(_as_matrix(avg_range).flat())
    vals = [float(flat[i]) for i, keep in enumerate(mask)
            if keep and i < len(flat)
            and isinstance(flat[i], (int, float)) and not isinstance(flat[i], bool)]
    if not vals:
        raise ExcelError(DIV0)
    return math.fsum(vals) / len(vals)


def _extreme_ifs(rng, args, pick):
    mask = _criteria_mask(_pairs(args))
    flat = list(_as_matrix(rng).flat())
    vals = [float(flat[i]) for i, keep in enumerate(mask)
            if keep and i < len(flat)
            and isinstance(flat[i], (int, float)) and not isinstance(flat[i], bool)]
    return pick(vals) if vals else 0.0


def f_maxifs(rng, *args):
    return _extreme_ifs(rng, args, max)


def f_minifs(rng, *args):
    return _extreme_ifs(rng, args, min)


def f_sumproduct(*args):
    arrays = [list(_as_matrix(a).flat()) for a in args]
    if not arrays:
        raise ExcelError(VALUE)
    n = len(arrays[0])
    if any(len(a) != n for a in arrays):
        raise ExcelError(VALUE)
    total = 0.0
    for i in range(n):
        product = 1.0
        for a in arrays:
            v = a[i]
            if isinstance(v, ExcelError):
                raise v
            product *= 0.0 if isinstance(v, (str, Blank)) else to_number(v)
        total += product
    return total


# --------------------------------------------------------------------- logic

def f_ifs(*args):
    for i in range(0, len(args) - 1, 2):
        if to_bool(args[i]):
            return args[i + 1]
    raise ExcelError(NA)


def f_switch(value, *args):
    default = args[-1] if len(args) % 2 == 1 else None
    for i in range(0, len(args) - 1, 2):
        if compare(value, args[i]) == 0:
            return args[i + 1]
    if default is None:
        raise ExcelError(NA)
    return default


def f_choose(index, *options):
    i = int(to_number(index))
    if i < 1 or i > len(options):
        raise ExcelError(VALUE)
    return options[i - 1]


def f_xlookup(needle, lookup_array, return_array, if_not_found=None, match_mode=0.0, _search=1.0):
    look = list(_as_matrix(lookup_array).flat())
    give = list(_as_matrix(return_array).flat())
    mode = int(to_number(match_mode)) if not isinstance(match_mode, Blank) else 0
    best = None
    for i, v in enumerate(look):
        try:
            c = compare(v, needle)
        except ExcelError:
            continue
        if c == 0:
            return give[i] if i < len(give) else ExcelError(REF_CODE)
        if mode == -1 and c < 0 and (best is None or compare(look[best], v) < 0):
            best = i
        if mode == 1 and c > 0 and (best is None or compare(v, look[best]) < 0):
            best = i
    if best is not None and best < len(give):
        return give[best]
    if if_not_found is not None and not isinstance(if_not_found, Blank):
        return if_not_found
    raise ExcelError(NA)


# ---------------------------------------------------------------------- dates
# Dates are numbers to Excel, so every one of these goes through to_serial /
# from_serial rather than carrying datetime objects around.

def f_today():
    return float(int(to_serial(_dt.date.today())))


def f_now():
    return to_serial(_dt.datetime.now())


def f_date(year, month, day):
    y, m, d = int(to_number(year)), int(to_number(month)), int(to_number(day))
    y += (m - 1) // 12                       # Excel rolls month overflow into the year
    m = (m - 1) % 12 + 1
    try:
        base = _dt.date(y, m, 1)
    except ValueError as e:
        raise ExcelError(NUM) from e
    return to_serial(base) + (d - 1)


def _as_date(v) -> _dt.datetime:
    return from_serial(to_number(v))


def f_year(v):
    return float(_as_date(v).year)


def f_month(v):
    return float(_as_date(v).month)


def f_day(v):
    return float(_as_date(v).day)


def f_hour(v):
    return float(_as_date(v).hour)


def f_minute(v):
    return float(_as_date(v).minute)


def f_weekday(v, return_type=1.0):
    dow = _as_date(v).weekday()              # Monday = 0
    t = int(to_number(return_type))
    if t == 2:
        return float(dow + 1)                # Monday = 1
    if t == 3:
        return float(dow)
    return float((dow + 1) % 7 + 1)          # default: Sunday = 1


def _add_months(d: _dt.datetime, months: int) -> _dt.date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    last = _MONTH_DAYS(y, m)
    return _dt.date(y, m, min(d.day, last))


def _MONTH_DAYS(y: int, m: int) -> int:
    import calendar
    return calendar.monthrange(y, m)[1]


def f_edate(start, months):
    return to_serial(_add_months(_as_date(start), int(to_number(months))))


def f_eomonth(start, months):
    d = _add_months(_as_date(start).replace(day=1), int(to_number(months)))
    return to_serial(_dt.date(d.year, d.month, _MONTH_DAYS(d.year, d.month)))


def f_days(end, start):
    return to_number(end) - to_number(start)


def f_datedif(start, end, unit):
    a, b = _as_date(start), _as_date(end)
    u = to_text(unit).upper()
    if b < a:
        raise ExcelError(NUM)
    if u == "D":
        return float((b - a).days)
    months = (b.year - a.year) * 12 + (b.month - a.month) - (1 if b.day < a.day else 0)
    if u == "M":
        return float(months)
    if u == "Y":
        return float(months // 12)
    raise ExcelError(NUM)


# ------------------------------------------------------------------- finance

def f_npv(rate, *values):
    r = to_number(rate)
    flat = [v for a in values for v in (a.flat() if isinstance(a, Matrix) else [a])]
    total = 0.0
    i = 1
    for v in flat:
        if isinstance(v, ExcelError):
            raise v
        if isinstance(v, (str, Blank)):
            continue
        total += to_number(v) / (1 + r) ** i
        i += 1
    return total


def f_irr(values, guess=0.1):
    flat = [to_number(v) for a in ([values] if not isinstance(values, Matrix) else [values])
            for v in (a.flat() if isinstance(a, Matrix) else [a])
            if not isinstance(v, (str, Blank))]
    if not flat or all(v >= 0 for v in flat) or all(v <= 0 for v in flat):
        raise ExcelError(NUM)

    def npv(r: float) -> float:
        return sum(v / (1 + r) ** i for i, v in enumerate(flat))

    lo, hi = -0.9999, 10.0
    f_lo, f_hi = npv(lo), npv(hi)
    if f_lo * f_hi > 0:
        raise ExcelError(NUM)
    for _ in range(200):                      # bisection: slower than Newton, always converges
        mid = (lo + hi) / 2
        f_mid = npv(mid)
        if abs(f_mid) < 1e-12:
            return mid
        if f_lo * f_mid < 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def _annuity(rate: float, nper: float, when: int) -> float:
    return (1 + rate) ** nper, (1 + rate * when)


def f_pmt(rate, nper, pv, fv=0.0, when=0.0):
    r, n = to_number(rate), to_number(nper)
    present, future, w = to_number(pv), to_number(fv), int(to_number(when))
    if n == 0:
        raise ExcelError(NUM)
    if r == 0:
        return -(present + future) / n
    growth = (1 + r) ** n
    return -(present * growth + future) * r / ((1 + r * w) * (growth - 1))


def f_pv(rate, nper, pmt, fv=0.0, when=0.0):
    r, n = to_number(rate), to_number(nper)
    payment, future, w = to_number(pmt), to_number(fv), int(to_number(when))
    if r == 0:
        return -(future + payment * n)
    growth = (1 + r) ** n
    return -(future + payment * (1 + r * w) * (growth - 1) / r) / growth


def f_fv(rate, nper, pmt, pv=0.0, when=0.0):
    r, n = to_number(rate), to_number(nper)
    payment, present, w = to_number(pmt), to_number(pv), int(to_number(when))
    if r == 0:
        return -(present + payment * n)
    growth = (1 + r) ** n
    return -(present * growth + payment * (1 + r * w) * (growth - 1) / r)


# ---------------------------------------------------------------------- maths

def f_mod(a, b):
    x, y = to_number(a), to_number(b)
    if y == 0:
        raise ExcelError(DIV0)
    return x - y * math.floor(x / y)          # Excel's MOD takes the divisor's sign


def f_ceiling(x, significance=1.0):
    v, s = to_number(x), to_number(significance)
    if s == 0:
        return 0.0
    if v > 0 and s < 0:
        raise ExcelError(NUM)
    return math.ceil(v / s - 1e-12) * s


def f_floor(x, significance=1.0):
    v, s = to_number(x), to_number(significance)
    if s == 0:
        raise ExcelError(DIV0)
    if v > 0 and s < 0:
        raise ExcelError(NUM)
    return math.floor(v / s + 1e-12) * s


def f_sign(x):
    v = to_number(x)
    return 0.0 if v == 0 else (1.0 if v > 0 else -1.0)


def f_median(*args):
    nums = sorted(_numbers_in(args))
    if not nums:
        raise ExcelError(NUM)
    mid = len(nums) // 2
    return nums[mid] if len(nums) % 2 else (nums[mid - 1] + nums[mid]) / 2


def f_stdev(*args):
    nums = _numbers_in(args)
    if len(nums) < 2:
        raise ExcelError(DIV0)
    mean = math.fsum(nums) / len(nums)
    return math.sqrt(math.fsum((v - mean) ** 2 for v in nums) / (len(nums) - 1))


def f_large(rng, k):
    nums = sorted(_numbers_in([rng]), reverse=True)
    i = int(to_number(k))
    if i < 1 or i > len(nums):
        raise ExcelError(NUM)
    return nums[i - 1]


def f_small(rng, k):
    nums = sorted(_numbers_in([rng]))
    i = int(to_number(k))
    if i < 1 or i > len(nums):
        raise ExcelError(NUM)
    return nums[i - 1]


# ----------------------------------------------------------------------- text

def f_textjoin(delimiter, ignore_empty, *args):
    sep = to_text(delimiter)
    skip = to_bool(ignore_empty)
    parts = []
    for a in args:
        for v in (a.flat() if isinstance(a, Matrix) else [a]):
            if isinstance(v, ExcelError):
                raise v
            t = to_text(v)
            if skip and (isinstance(v, Blank) or t == ""):
                continue
            parts.append(t)
    return sep.join(parts)


def f_substitute(text, old, new, instance=None):
    s, o, n = to_text(text), to_text(old), to_text(new)
    if o == "":
        return s
    if instance is None or isinstance(instance, Blank):
        return s.replace(o, n)
    i = int(to_number(instance))
    if i < 1:
        raise ExcelError(VALUE)
    parts = s.split(o)
    if len(parts) <= i:
        return s
    return o.join(parts[:i]) + n + o.join(parts[i:])


def f_replace(text, start, count, new):
    s = to_text(text)
    i, n = int(to_number(start)), int(to_number(count))
    if i < 1:
        raise ExcelError(VALUE)
    return s[: i - 1] + to_text(new) + s[i - 1 + n:]


def f_find(needle, haystack, start=1.0):
    i = to_text(haystack).find(to_text(needle), int(to_number(start)) - 1)
    if i < 0:
        raise ExcelError(VALUE)
    return float(i + 1)


def f_search(needle, haystack, start=1.0):
    i = to_text(haystack).upper().find(to_text(needle).upper(), int(to_number(start)) - 1)
    if i < 0:
        raise ExcelError(VALUE)
    return float(i + 1)


def f_proper(v):
    return to_text(v).title()


def f_rept(v, n):
    count = int(to_number(n))
    if count < 0:
        raise ExcelError(VALUE)
    return to_text(v) * count


def f_isna(v):
    return isinstance(v, ExcelError) and v.code == NA


def f_iferror_na(v):
    return isinstance(v, ExcelError) and v.code != NA


def f_na():
    raise ExcelError(NA)


def f_n(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    return 0.0


def f_hlookup(needle, table, row_index, approximate=True):
    m = _as_matrix(table)
    ri = int(to_number(row_index))
    if ri < 1 or ri > len(m):
        raise ExcelError(VALUE)
    header = m[0]
    approx = to_bool(approximate) if not isinstance(approximate, bool) else approximate
    best = None
    for j, v in enumerate(header):
        try:
            c = compare(v, needle)
        except ExcelError:
            continue
        if c == 0:
            return m[ri - 1][j]
        if approx and c < 0:
            best = j
    if best is None:
        raise ExcelError(NA)
    return m[ri - 1][best]


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
    "VLOOKUP": f_vlookup, "HLOOKUP": f_hlookup, "INDEX": f_index, "MATCH": f_match,
    "XLOOKUP": f_xlookup, "CHOOSE": f_choose,
    # multi-criteria aggregates
    "SUMIFS": f_sumifs, "COUNTIFS": f_countifs, "AVERAGEIFS": f_averageifs,
    "MAXIFS": f_maxifs, "MINIFS": f_minifs, "SUMPRODUCT": f_sumproduct,
    # logic
    "IFS": f_ifs, "SWITCH": f_switch, "ISNA": f_isna, "NA": f_na, "N": f_n,
    # dates
    "TODAY": f_today, "NOW": f_now, "DATE": f_date, "YEAR": f_year, "MONTH": f_month,
    "DAY": f_day, "HOUR": f_hour, "MINUTE": f_minute, "WEEKDAY": f_weekday,
    "EDATE": f_edate, "EOMONTH": f_eomonth, "DAYS": f_days, "DATEDIF": f_datedif,
    # finance
    "NPV": f_npv, "IRR": f_irr, "PMT": f_pmt, "PV": f_pv, "FV": f_fv,
    # maths
    "MOD": f_mod, "CEILING": f_ceiling, "FLOOR": f_floor, "SIGN": f_sign,
    "MEDIAN": f_median, "STDEV": f_stdev, "STDEV.S": f_stdev,
    "LARGE": f_large, "SMALL": f_small,
    # text
    "TEXTJOIN": f_textjoin, "SUBSTITUTE": f_substitute, "REPLACE": f_replace,
    "FIND": f_find, "SEARCH": f_search, "PROPER": f_proper, "REPT": f_rept,
}

#: Functions whose result depends on more than the workbook (recalculated every open).
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"}

#: Functions Gridlint can name but cannot evaluate, so a formula using one is
#: reported as unmodelled rather than silently given a wrong value.
UNSUPPORTED = {"OFFSET", "INDIRECT", "RAND", "RANDBETWEEN", "CELL", "INFO", "LET",
               "LAMBDA", "FILTER", "SORT", "UNIQUE", "SEQUENCE", "TEXTSPLIT"}
