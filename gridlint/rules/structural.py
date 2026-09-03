"""Structural defects: ranges that miss data, formulas that break a pattern, totals that double-count.

These are the defects that cost money, because the spreadsheet keeps working
and shows a plausible number. Nothing here asks a language model anything.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Any, Iterable

from ..formula.ast_nodes import FuncCall, Node, RangeRef, Ref, walk
from ..formula.parser import index_to_col
from ..formula.values import BLANK, ExcelError
from ..workbook import Cell, Workbook
from .base import CRITICAL, WARNING, Finding, Fix, addr, rule, shape_of, skeleton_of, to_r1c1

#: Aggregates whose range is expected to cover a whole block of data.
BLOCK_AGGREGATES = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "MEDIAN", "STDEV"}


def _is_numeric(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _cell_value(wb: Workbook, computed: dict, sheet: str, col: int, row: int) -> Any:
    key = (sheet, col, row)
    if key in computed:
        return computed[key]
    c = wb.get(sheet, col, row)
    if c is None:
        return BLANK
    return c.static if c.static is not None else (c.cached if c.cached is not None else BLANK)


def _looks_like_data(wb: Workbook, computed: dict, sheet: str, col: int, row: int) -> bool:
    c = wb.get(sheet, col, row)
    if c is None:
        return False
    return _is_numeric(_cell_value(wb, computed, sheet, col, row))


def _row_label(wb: Workbook, sheet: str, row: int, before_col: int) -> str | None:
    """The nearest text label to the left, which is what makes a row a real line item."""
    for c in range(1, min(before_col, 6)):
        cell = wb.get(sheet, c, row)
        if cell is not None and isinstance(cell.static, str) and cell.static.strip():
            return cell.static.strip()
    return None


@rule("R001", "Aggregate range misses adjacent data",
      "A SUM or AVERAGE that stops one row short keeps working and shows a smaller, plausible number.",
      CRITICAL)
def range_omission(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        for node in walk(cell.ast):
            if not isinstance(node, FuncCall) or node.name not in BLOCK_AGGREGATES:
                continue
            for arg in node.args:
                if not isinstance(arg, RangeRef):
                    continue
                sheet = arg.sheet or cell.sheet
                vertical = arg.col1 == arg.col2
                horizontal = arg.row1 == arg.row2
                if not (vertical or horizontal) or arg.n_cells < 2:
                    continue
                for missed, side in _adjacent_data(wb, computed, sheet, arg, cell, vertical):
                    mc, mr = missed
                    label = _row_label(wb, sheet, mr, arg.col1) if vertical else None
                    value = _cell_value(wb, computed, sheet, mc, mr)
                    new_range = _extend(arg, side, vertical)
                    new_formula = cell.formula.replace(arg.raw, new_range, 1) if cell.formula else None
                    if not new_formula or new_formula == cell.formula:
                        continue
                    what = f'"{label}"' if label else addr(sheet, mc, mr)
                    yield Finding(
                        rule="R001",
                        title="Total leaves out a row of data",
                        severity=CRITICAL,
                        cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                        formula=cell.formula,
                        detail=(f"{node.name} covers {arg.raw} but {what} sits right "
                                f"{'below' if side == 'after' else 'above'} it and is left out."),
                        evidence={"range": arg.raw, "missing_cell": addr(sheet, mc, mr),
                                  "missing_label": label, "missing_value": _plain(value),
                                  "function": node.name, "side": side},
                        fix=Fix.one(cell.ref, new_formula,
                                    label=f"Extend {arg.raw} to {new_range}",
                                    old_formula=cell.formula),
                        confidence=0.9 if label else 0.7,
                    )


def _adjacent_data(wb: Workbook, computed: dict, sheet: str, rng: RangeRef, cell: Cell, vertical: bool):
    """Yield ((col,row), side) for data cells immediately outside the range."""
    own = (cell.sheet, cell.col, cell.row)
    if vertical:
        col = rng.col1
        after = (col, rng.row2 + 1)
        before = (col, rng.row1 - 1)
    else:
        row = rng.row1
        after = (rng.col2 + 1, row)
        before = (rng.col1 - 1, row)

    for pos, side in ((after, "after"), (before, "before")):
        c, r = pos
        if c < 1 or r < 1:
            continue
        if (sheet, c, r) == own:
            continue
        if not _looks_like_data(wb, computed, sheet, c, r):
            continue
        neighbour = wb.get(sheet, c, r)
        if neighbour is None:
            continue
        # The neighbour must belong to the same block: same shape as the range's own
        # members (both plain numbers, or both formulas of the same shape).
        if not _same_block(wb, sheet, rng, neighbour, vertical):
            continue
        # A cell that already aggregates this very range is a total, not a data row.
        if _aggregates_range(neighbour, rng, sheet):
            continue
        yield pos, side


def _same_block(wb: Workbook, sheet: str, rng: RangeRef, neighbour: Cell, vertical: bool) -> bool:
    members = []
    if vertical:
        for r in range(rng.row1, min(rng.row2, rng.row1 + 30) + 1):
            m = wb.get(sheet, rng.col1, r)
            if m is not None:
                members.append(m)
    else:
        for c in range(rng.col1, min(rng.col2, rng.col1 + 30) + 1):
            m = wb.get(sheet, c, rng.row1)
            if m is not None:
                members.append(m)
    if not members:
        return False
    n_formula = sum(1 for m in members if m.formula)
    if neighbour.formula:
        if n_formula == 0:
            return False
        # Compare skeletons, not exact shapes: sibling rows legitimately differ in
        # their absolute references (=C11*(1+$B$11) vs =C12*(1+$B$12)) while doing
        # exactly the same kind of calculation.
        skeletons = {skeleton_of(m.ast) for m in members if m.formula}
        return skeleton_of(neighbour.ast) in skeletons
    return any(m.formula is None for m in members)


def _aggregates_range(cell: Cell, rng: RangeRef, sheet: str) -> bool:
    if cell.ast is None:
        return False
    for node in walk(cell.ast):
        if isinstance(node, RangeRef):
            if (node.sheet or cell.sheet) == sheet and node.col1 == rng.col1 and node.row1 == rng.row1:
                return True
    return False


def _extend(rng: RangeRef, side: str, vertical: bool) -> str:
    sheet = f"{rng.sheet}!" if rng.sheet else ""
    if vertical:
        c = index_to_col(rng.col1)
        r1 = rng.row1 - 1 if side == "before" else rng.row1
        r2 = rng.row2 + 1 if side == "after" else rng.row2
        return f"{sheet}{c}{r1}:{c}{r2}"
    r = rng.row1
    c1 = rng.col1 - 1 if side == "before" else rng.col1
    c2 = rng.col2 + 1 if side == "after" else rng.col2
    return f"{sheet}{index_to_col(c1)}{r}:{index_to_col(c2)}{r}"


@rule("R002", "Formula breaks the pattern of its neighbours",
      "One cell in a row of copied formulas was edited by hand; every other cell still agrees.",
      CRITICAL)
def inconsistent_formula(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for sheet in wb.sheet_list:
        yield from _scan_runs(wb, sheet, horizontal=True)
        yield from _scan_runs(wb, sheet, horizontal=False)


def _scan_runs(wb: Workbook, sheet, horizontal: bool) -> Iterable[Finding]:
    buckets: dict[int, list[Cell]] = defaultdict(list)
    for (col, row), cell in sheet.cells.items():
        if cell.formula and cell.ast is not None:
            buckets[row if horizontal else col].append(cell)

    for _line, cells in buckets.items():
        cells.sort(key=lambda c: c.col if horizontal else c.row)
        for run in _contiguous(cells, horizontal):
            if len(run) < 4:
                continue
            sigs = [to_r1c1(c.ast, c.col, c.row) for c in run]
            counts = defaultdict(int)
            for s in sigs:
                counts[s] += 1
            majority, n = max(counts.items(), key=lambda kv: kv[1])
            if n < len(run) - 2 or n / len(run) < 0.7:
                continue
            for cell, sig in zip(run, sigs):
                if sig == majority:
                    continue
                twin = next(c for c, s in zip(run, sigs) if s == majority)
                repaired = _translate(twin, cell)
                axis = "row" if horizontal else "column"
                yield Finding(
                    rule="R002",
                    title="One formula differs from the rest of the line",
                    severity=CRITICAL,
                    cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                    formula=cell.formula,
                    detail=(f"{n} of {len(run)} formulas in this {axis} follow one pattern; "
                            f"{cell.addr} does something different."),
                    evidence={"pattern": majority, "this": sig, "run": len(run),
                              "agreeing": n, "example": twin.ref, "example_formula": twin.formula},
                    fix=(Fix.one(cell.ref, repaired,
                                 label=f"Match the pattern used by {twin.addr}",
                                 old_formula=cell.formula) if repaired else None),
                    related=[twin.ref],
                    confidence=min(0.95, 0.55 + 0.4 * (n / len(run))),
                )


def _contiguous(cells: list[Cell], horizontal: bool) -> Iterable[list[Cell]]:
    run: list[Cell] = []
    prev = None
    for c in cells:
        pos = c.col if horizontal else c.row
        if prev is not None and pos != prev + 1:
            if len(run) >= 4:
                yield run
            run = []
        run.append(c)
        prev = pos
    if len(run) >= 4:
        yield run


def _translate(source: Cell, target: Cell) -> str | None:
    """Rewrite `source`'s formula as if it had been copied into `target`'s cell."""
    try:
        from openpyxl.formula.translate import Translator
        return Translator(source.formula, origin=source.addr).translate_formula(target.addr)
    except Exception:
        return None


@rule("R009", "Total counts another total",
      "A range that contains a subtotal adds the same numbers twice.",
      CRITICAL)
def double_counting(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    # Only cells that themselves contain a SUM can be double counted, and there
    # are few of them. Indexing those once turns a scan of every cell in every
    # range into a lookup, which is the difference between seconds and
    # milliseconds on a real model.
    summing: dict[tuple[str, int], list[Cell]] = defaultdict(list)
    for c in wb.formula_cells():
        if c.ast is not None and any(isinstance(n, FuncCall) and n.name in ("SUM", "SUBTOTAL")
                                     for n in walk(c.ast)):
            summing[(c.sheet, c.row)].append(c)

    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        for node in walk(cell.ast):
            if not isinstance(node, FuncCall) or node.name not in ("SUM", "PRODUCT"):
                continue
            for arg in node.args:
                if not isinstance(arg, RangeRef):
                    continue
                sheet = arg.sheet or cell.sheet
                rows = range(arg.row1, min(arg.row2, arg.row1 + 100_000) + 1)
                for inner in (c for r in rows for c in summing.get((sheet, r), ())):
                    if not arg.col1 <= inner.col <= arg.col2:
                        continue
                    if (inner.sheet, inner.col, inner.row) == (cell.sheet, cell.col, cell.row):
                        continue
                    if True:
                        if not _sums_within(inner, arg, sheet):
                            continue
                        yield Finding(
                            rule="R009",
                            title="Total includes a subtotal from the same block",
                            severity=CRITICAL,
                            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                            formula=cell.formula,
                            detail=(f"{arg.raw} contains {inner.addr}, which already adds up cells "
                                    f"inside that same range, so those numbers are counted twice."),
                            evidence={"range": arg.raw, "subtotal": inner.ref,
                                      "subtotal_formula": inner.formula},
                            related=[inner.ref],
                            confidence=0.85,
                        )


def _sums_within(inner: Cell, outer: RangeRef, sheet: str) -> bool:
    if inner.ast is None:
        return False
    for node in walk(inner.ast):
        if isinstance(node, FuncCall) and node.name in ("SUM", "SUBTOTAL"):
            for arg in node.args:
                if not isinstance(arg, RangeRef):
                    continue
                if (arg.sheet or inner.sheet) != sheet:
                    continue
                if (arg.col1 >= outer.col1 and arg.col2 <= outer.col2
                        and arg.row1 >= outer.row1 and arg.row2 <= outer.row2):
                    return True
    return False


def _plain(v: Any) -> Any:
    if isinstance(v, ExcelError):
        return v.code
    if v is BLANK:
        return None
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    return v
