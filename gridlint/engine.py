"""The calculation engine: dependency graph, topological recalculation, self-check.

The engine is what separates Gridlint from a linter that only pattern-matches
text. Because it can recompute the workbook, it can answer the two questions a
reviewer actually cares about:

  * "how much money does this defect move?"  -> recalculate with the fix, diff
  * "can I trust your recalculation?"        -> compare against the values the
                                                spreadsheet app itself cached
"""
from __future__ import annotations

import datetime as _dt
import math
from dataclasses import dataclass, field
from typing import Any, Iterable

from .formula.ast_nodes import ExternalRef, FuncCall, Node, RangeRef, Ref, TableRef, walk
from .formula.evaluator import CIRC, evaluate
from .formula.functions import VOLATILE
from .formula.parser import parse_formula, parse_ref, parse_range
from .formula.values import BLANK, ExcelError, to_serial
from .workbook import Cell, Workbook, a1

CellKey = tuple[str, int, int]           # (sheet, col, row)

#: Ranges wider than this are treated as opaque for precedent tracking, to keep
#: whole-column references (A:A = 1,048,576 cells) from exploding the graph.
MAX_EXPAND = 50_000


@dataclass
class Precedents:
    cells: set[CellKey] = field(default_factory=set)
    ranges: list[tuple[str, int, int, int, int]] = field(default_factory=list)
    volatile: bool = False
    unresolved: bool = False


def _clamp(wb: Workbook, sheet: str, node: RangeRef) -> tuple[int, int, int, int]:
    s = wb.sheets.get(sheet)
    max_col = s.max_col if s else node.col2
    max_row = s.max_row if s else node.row2
    return (node.col1, node.row1,
            min(node.col2, max(max_col, node.col1)),
            min(node.row2, max(max_row, node.row1)))


def table_range_of(wb: Workbook, table: str, column: str | None) -> RangeRef | None:
    """Turn Table1[Amount] into the range it actually covers."""
    info = wb.tables.get(table.lower())
    if info is None:
        return None
    try:
        c1, r1, c2, r2 = info.column_range(column)
    except KeyError:
        return None
    if r2 < r1:
        return None
    return RangeRef(info.sheet, c1, r1, c2, r2, raw=f"{table}[{column or ''}]")


def precedents_of(wb: Workbook, cell: Cell) -> Precedents:
    """Cells this formula reads. Ranges are expanded within the sheet's used area."""
    p = Precedents()
    if cell.ast is None:
        p.unresolved = True
        return p
    for node in walk(cell.ast):
        if isinstance(node, Ref):
            p.cells.add((node.sheet or cell.sheet, node.col, node.row))
        elif isinstance(node, RangeRef):
            sheet = node.sheet or cell.sheet
            c1, r1, c2, r2 = _clamp(wb, sheet, node)
            p.ranges.append((sheet, c1, r1, c2, r2))
            if (c2 - c1 + 1) * (r2 - r1 + 1) <= MAX_EXPAND:
                for c in range(c1, c2 + 1):
                    for r in range(r1, r2 + 1):
                        p.cells.add((sheet, c, r))
            else:
                p.unresolved = True
        elif isinstance(node, TableRef):
            resolved = table_range_of(wb, node.table, node.column)
            if resolved is None:
                p.unresolved = True
            else:
                c1, r1, c2, r2 = _clamp(wb, resolved.sheet or cell.sheet, resolved)
                p.ranges.append((resolved.sheet or cell.sheet, c1, r1, c2, r2))
                if (c2 - c1 + 1) * (r2 - r1 + 1) <= MAX_EXPAND:
                    for c in range(c1, c2 + 1):
                        for r in range(r1, r2 + 1):
                            p.cells.add((resolved.sheet or cell.sheet, c, r))
        elif isinstance(node, ExternalRef):
            p.unresolved = True
        elif isinstance(node, FuncCall) and node.name in VOLATILE:
            p.volatile = True
    return p


class _Resolver:
    """Feeds the evaluator values from an override map, then computed, then static."""

    def __init__(self, wb: Workbook, computed: dict[CellKey, Any], overrides: dict[CellKey, Any] | None = None):
        self.wb = wb
        self.computed = computed
        self.overrides = overrides or {}

    def cell(self, sheet: str | None, col: int, row: int) -> Any:
        key = (sheet or "", col, row)
        if key in self.overrides:
            return self.overrides[key]
        if key in self.computed:
            return self.computed[key]
        c = self.wb.get(key[0], col, row)
        if c is None:
            return BLANK
        if c.formula:
            return c.cached if c.cached is not None else BLANK
        return c.static if c.static is not None else BLANK

    def used_bounds(self, sheet: str | None) -> tuple[int, int]:
        s = self.wb.sheets.get(sheet or "")
        return (s.max_col, s.max_row) if s else (0, 0)

    def sheet_exists(self, sheet: str) -> bool:
        return sheet in self.wb.sheets

    def table_range(self, table: str, column: str | None) -> RangeRef | None:
        return table_range_of(self.wb, table, column)

    def defined_name(self, name: str) -> Any:
        target = self.wb.defined_names.get(name.upper())
        if not target:
            return None
        try:
            return parse_formula(target if target.startswith("=") else "=" + target.replace("$", "$"))
        except Exception:
            return None


@dataclass
class Graph:
    order: list[CellKey]                                   # topological, dependencies first
    precedents: dict[CellKey, Precedents]
    dependents: dict[CellKey, set[CellKey]]
    cycles: list[list[CellKey]]

    def downstream(self, seeds: Iterable[CellKey], limit: int = 100_000) -> set[CellKey]:
        """Every cell whose value can change when `seeds` change."""
        seen: set[CellKey] = set()
        stack = list(seeds)
        while stack and len(seen) < limit:
            k = stack.pop()
            for d in self.dependents.get(k, ()):
                if d not in seen:
                    seen.add(d)
                    stack.append(d)
        return seen


def build_graph(wb: Workbook) -> Graph:
    precedents: dict[CellKey, Precedents] = {}
    dependents: dict[CellKey, set[CellKey]] = {}
    formula_keys: list[CellKey] = []

    for cell in wb.formula_cells():
        key = (cell.sheet, cell.col, cell.row)
        formula_keys.append(key)
        p = precedents_of(wb, cell)
        precedents[key] = p
        for dep in p.cells:
            dependents.setdefault(dep, set()).add(key)

    formula_set = set(formula_keys)
    indeg = {k: sum(1 for d in precedents[k].cells if d in formula_set and d != k) for k in formula_keys}
    ready = sorted([k for k, d in indeg.items() if d == 0])
    order: list[CellKey] = []
    while ready:
        k = ready.pop()
        order.append(k)
        for dep in sorted(dependents.get(k, ())):
            if dep in indeg:
                indeg[dep] -= 1
                if indeg[dep] == 0:
                    ready.append(dep)

    ordered = set(order)                    # build once: this was O(n^2) inside the comprehension
    stuck = [k for k in formula_keys if k not in ordered]
    cycles = _find_cycles(stuck, precedents, formula_set) if stuck else []
    order.extend(stuck)                                    # evaluated last, will report #CIRC!
    return Graph(order=order, precedents=precedents, dependents=dependents, cycles=cycles)


def _find_cycles(stuck: list[CellKey], precedents: dict[CellKey, Precedents],
                 formula_set: set[CellKey]) -> list[list[CellKey]]:
    """Tarjan's strongly connected components, restricted to cells that never became ready."""
    index: dict[CellKey, int] = {}
    low: dict[CellKey, int] = {}
    on_stack: set[CellKey] = set()
    stack: list[CellKey] = []
    out: list[list[CellKey]] = []
    counter = [0]
    nodes = set(stuck)

    def strongconnect(v: CellKey) -> None:
        work = [(v, iter(sorted(p for p in precedents.get(v, Precedents()).cells if p in nodes)))]
        index[v] = low[v] = counter[0]
        counter[0] += 1
        stack.append(v)
        on_stack.add(v)
        while work:
            node, it = work[-1]
            advanced = False
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter[0]
                    counter[0] += 1
                    stack.append(w)
                    on_stack.add(w)
                    work.append((w, iter(sorted(p for p in precedents.get(w, Precedents()).cells if p in nodes))))
                    advanced = True
                    break
                if w in on_stack:
                    low[node] = min(low[node], index[w])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
            if low[node] == index[node]:
                comp = []
                while True:
                    w = stack.pop()
                    on_stack.discard(w)
                    comp.append(w)
                    if w == node:
                        break
                if len(comp) > 1 or node in precedents.get(node, Precedents()).cells:
                    out.append(sorted(comp))

    for v in sorted(nodes):
        if v not in index:
            strongconnect(v)
    return out


def recalc(wb: Workbook, graph: Graph | None = None,
           overrides: dict[CellKey, Any] | None = None,
           formula_overrides: dict[CellKey, str] | None = None) -> dict[CellKey, Any]:
    """Recompute every formula cell. Returns {cell key: value}.

    `overrides` replaces input values; `formula_overrides` replaces formulas,
    which is how a candidate fix is tested before anyone accepts it.
    """
    if formula_overrides:
        graph = build_graph(_with_formulas(wb, formula_overrides))
        wb = _with_formulas(wb, formula_overrides)
    elif graph is None:
        graph = build_graph(wb)

    computed: dict[CellKey, Any] = {}
    resolver = _Resolver(wb, computed, overrides)
    cyclic = {k for comp in graph.cycles for k in comp}
    for key in graph.order:
        cell = wb.get(*key)
        if cell is None or cell.ast is None:
            computed[key] = ExcelError("#NAME?") if cell and cell.parse_error else BLANK
            continue
        if key in cyclic:
            computed[key] = ExcelError(CIRC)
            continue
        computed[key] = evaluate(cell.ast, cell.sheet, resolver)
    return computed


def _with_formulas(wb: Workbook, formula_overrides: dict[CellKey, str]) -> Workbook:
    """A shallow copy of the workbook with some formulas replaced."""
    import copy

    clone = Workbook(path=wb.path, defined_names=dict(wb.defined_names))
    for name, sheet in wb.sheets.items():
        s = type(sheet)(name=sheet.name, index=sheet.index, max_col=sheet.max_col,
                        max_row=sheet.max_row, hidden=sheet.hidden)
        s.cells = dict(sheet.cells)
        clone.sheets[name] = s
    for key, formula in formula_overrides.items():
        sheet_name, col, row = key
        s = clone.sheets.get(sheet_name)
        if s is None:
            continue
        old = s.cells.get((col, row))
        new = copy.copy(old) if old else Cell(sheet=sheet_name, col=col, row=row)
        new.formula = formula
        new.static = None
        new.parse_error = None
        try:
            new.ast = parse_formula(formula)
        except Exception as e:
            new.ast = None
            new.parse_error = str(e)
        s.cells[(col, row)] = new
        s.max_col = max(s.max_col, col)
        s.max_row = max(s.max_row, row)
    return clone


#: Excel stores doubles with about 15 significant digits. The tolerance has to
#: absorb that last-digit noise and nothing more: at a relative 1e-9, a million
#: dollars is compared to a tenth of a cent, so a genuinely stale number is
#: never mistaken for rounding.
NUMERIC_TOLERANCE = 1e-9


def values_equal(a: Any, b: Any, tol: float = NUMERIC_TOLERANCE) -> bool:
    if isinstance(a, ExcelError) or isinstance(b, ExcelError):
        return isinstance(a, ExcelError) and isinstance(b, ExcelError) and a.code == b.code
    if a is None or b is None:
        return a is None and b is None
    # A date-formatted cell comes back from the file as a datetime while the
    # engine computes the serial number underneath it. They are the same value.
    if isinstance(a, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        a = to_serial(a)
    if isinstance(b, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        b = to_serial(b)
    if isinstance(a, bool) or isinstance(b, bool):
        return bool(a) == bool(b) and isinstance(a, bool) == isinstance(b, bool)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        fa, fb = float(a), float(b)
        if math.isnan(fa) or math.isnan(fb):
            return False
        scale = max(1.0, abs(fa), abs(fb))
        return abs(fa - fb) <= tol * scale
    if a is BLANK:
        return b is BLANK or b is None or b == "" or b == 0
    if b is BLANK:
        return a is BLANK or a is None or a == "" or a == 0
    return str(a) == str(b)


@dataclass
class TraceStep:
    """One hop in the chain from a headline number back to what drives it."""
    cell: str
    label: str | None
    formula: str | None
    value: Any
    depth: int
    is_input: bool           # a typed number rather than a computed one
    changed: bool = False    # this cell's value moves when the fix is applied


def trace(wb: Workbook, graph: "Graph", computed: dict[CellKey, Any], target: CellKey, *,
          max_depth: int = 6, max_steps: int = 60,
          changed: set[CellKey] | None = None) -> list[TraceStep]:
    """Walk backwards from a cell to the inputs it is built from.

    This is the question a reviewer actually asks -- "where does this number
    come from?" -- and it is only answerable because the dependency graph exists.
    Branches that do not move when the fix is applied are pruned first, so the
    chain shown is the chain that matters.
    """
    out: list[TraceStep] = []
    seen: set[CellKey] = set()
    queue: list[tuple[CellKey, int]] = [(target, 0)]
    while queue and len(out) < max_steps:
        key, depth = queue.pop(0)
        if key in seen or depth > max_depth:
            continue
        seen.add(key)
        cell = wb.get(*key)
        if cell is None:
            continue
        out.append(TraceStep(
            cell=f"{key[0]}!{a1(key[1], key[2])}",
            label=_label_left_of(wb, key[0], key[2], key[1]),
            formula=cell.formula,
            value=_plain(computed.get(key, cell.static if cell.static is not None else cell.cached)),
            depth=depth,
            is_input=cell.formula is None,
            changed=bool(changed and key in changed),
        ))
        prec = graph.precedents.get(key)
        if not prec:
            continue
        children = sorted(prec.cells)
        if changed:
            moving = [c for c in children if c in changed]
            children = moving or children[:4]
        for child in children[:8]:
            if child not in seen:
                queue.append((child, depth + 1))
    return out


def _label_left_of(wb: Workbook, sheet: str, row: int, before_col: int) -> str | None:
    for c in range(1, min(before_col, 8)):
        cell = wb.get(sheet, c, row)
        if cell is not None and isinstance(cell.static, str) and cell.static.strip():
            return cell.static.strip()
    return None


@dataclass
class SelfCheck:
    """How well the engine reproduced the values the spreadsheet app cached."""
    total: int = 0
    checked: int = 0
    matched: int = 0
    mismatches: list[dict[str, Any]] = field(default_factory=list)
    unsupported: list[dict[str, Any]] = field(default_factory=list)

    @property
    def agreement(self) -> float:
        return self.matched / self.checked if self.checked else 0.0

    @property
    def trustworthy(self) -> bool:
        """Only claim a money figure when the engine reproduced essentially everything."""
        return self.checked > 0 and self.agreement >= 0.995

    def summary(self) -> str:
        if not self.checked:
            return "no cached values to check against"
        return (f"reproduced {self.matched}/{self.checked} cached values "
                f"({self.agreement * 100:.1f}%)")


def self_check(wb: Workbook, computed: dict[CellKey, Any] | None = None,
               graph: Graph | None = None) -> SelfCheck:
    graph = graph or build_graph(wb)
    computed = computed if computed is not None else recalc(wb, graph)
    sc = SelfCheck()
    for cell in wb.formula_cells():
        sc.total += 1
        key = (cell.sheet, cell.col, cell.row)
        if cell.parse_error:
            sc.unsupported.append({"cell": cell.ref, "formula": cell.formula, "reason": cell.parse_error})
            continue
        if cell.cached is None:
            continue
        got = computed.get(key, BLANK)
        if isinstance(got, ExcelError) and got.code == "#NAME?" and not isinstance(cell.cached, ExcelError):
            reason = ("external workbook link" if cell.uses_external else
                      "table reference that could not be resolved" if cell.uses_table else
                      "unsupported function")
            sc.unsupported.append({"cell": cell.ref, "formula": cell.formula, "reason": reason})
            continue
        sc.checked += 1
        if values_equal(got, cell.cached):
            sc.matched += 1
        else:
            sc.mismatches.append({
                "cell": cell.ref, "formula": cell.formula,
                "expected": _plain(cell.cached), "got": _plain(got),
            })
    return sc


def _plain(v: Any) -> Any:
    if isinstance(v, ExcelError):
        return v.code
    if v is BLANK:
        return None
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    return v
