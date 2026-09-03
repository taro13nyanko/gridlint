"""Defects that make a model fragile rather than immediately wrong."""
from __future__ import annotations

from typing import Any, Iterable

from ..formula.ast_nodes import BinOp, FuncCall, Literal, Node, RangeRef, Ref, walk
from ..formula.values import BLANK, ExcelError
from ..workbook import Workbook
from .base import INFO, WARNING, Finding, addr, is_currency_format, rule

#: Literals that are almost always structural rather than an assumption.
BENIGN_LITERALS = {0.0, 1.0, 2.0, -1.0, 100.0, 12.0, 24.0, 60.0, 365.0, 7.0, 4.0, 1000.0}
#: Argument positions where a literal is the function's own machinery, not a business number.
STRUCTURAL_ARG = {
    "ROUND": {1}, "ROUNDUP": {1}, "ROUNDDOWN": {1}, "VLOOKUP": {2, 3}, "HLOOKUP": {2, 3},
    "INDEX": {1, 2}, "MATCH": {2}, "LEFT": {1}, "RIGHT": {1}, "MID": {1, 2}, "TEXT": {1},
    "IF": {1, 2}, "IFERROR": {1}, "IFNA": {1}, "POWER": {1}, "SUBTOTAL": {0},
}


@rule("R003", "A number is buried inside a formula",
      "An assumption written into a formula cannot be found, reviewed or changed by anyone else.", WARNING)
def hardcoded_constants(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cell in wb.formula_cells():
        if cell.ast is None or not cell.formula:
            continue
        literals = _business_literals(cell.ast)
        if not literals:
            continue
        has_refs = any(isinstance(n, (Ref, RangeRef)) for n in walk(cell.ast))
        if not has_refs:
            continue                                  # a pure constant expression is a value, not a hidden one
        shown = ", ".join(_fmt(v) for v in sorted(set(literals))[:4])
        yield Finding(
            rule="R003",
            title=f"Hidden assumption: {shown}",
            severity=WARNING,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=(f"The number {shown} is typed straight into this formula. "
                    "Put it in its own labelled cell so it can be found and changed in one place."),
            evidence={"literals": sorted(set(literals))[:8]},
            confidence=0.65,
        )


def _business_literals(node: Node) -> list[float]:
    out: list[float] = []

    def visit(n: Node, parent_fn: str | None, arg_index: int | None) -> None:
        if isinstance(n, Literal):
            v = n.value
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                return
            f = float(v)
            if f in BENIGN_LITERALS:
                return
            if parent_fn and arg_index in STRUCTURAL_ARG.get(parent_fn, set()):
                return
            out.append(f)
            return
        if isinstance(n, FuncCall):
            for i, a in enumerate(n.args):
                visit(a, n.name, i)
            return
        for f in ("operand", "left", "right"):
            child = getattr(n, f, None)
            if isinstance(child, Node):
                visit(child, parent_fn, arg_index)

    visit(node, None, None)
    return out


def _fmt(v: float) -> str:
    return str(int(v)) if v == int(v) else str(v)


@rule("R004", "Formula reads an empty cell",
      "An empty cell counts as zero, so a missing input looks like a real one.", WARNING)
def reference_to_blank(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        blanks: list[str] = []
        for node in walk(cell.ast):
            if not isinstance(node, Ref):
                continue
            sheet = node.sheet or cell.sheet
            target = wb.get(sheet, node.col, node.row)
            if target is None:
                blanks.append(addr(sheet, node.col, node.row))
        if not blanks:
            continue
        if _only_in_lookup_or_text(cell.ast):
            continue
        yield Finding(
            rule="R004",
            title=f"Reads empty cell {blanks[0]}",
            severity=WARNING,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=(f"{', '.join(blanks[:3])} {'is' if len(blanks) == 1 else 'are'} empty. "
                    "Excel treats an empty cell as zero, so this result looks valid even though an input is missing."),
            evidence={"blank_cells": blanks[:8]},
            confidence=0.7,
        )


def _only_in_lookup_or_text(node: Node) -> bool:
    for n in walk(node):
        if isinstance(n, FuncCall) and n.name in ("ISBLANK", "COUNTBLANK", "COUNTA", "IF", "IFERROR"):
            return True
    return False


@rule("R008", "A number in this range is stored as text",
      "Numbers stored as text are skipped by SUM, so the total is quietly short.", WARNING)
def numbers_stored_as_text(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    from ..rules.structural import BLOCK_AGGREGATES

    reported: set[str] = set()
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
                offenders = []
                for c in range(arg.col1, min(arg.col2, arg.col1 + 100) + 1):
                    for r in range(arg.row1, min(arg.row2, arg.row1 + 5000) + 1):
                        t = wb.get(sheet, c, r)
                        if t is None or t.formula or not isinstance(t.static, str):
                            continue
                        if _looks_numeric(t.static):
                            offenders.append((addr(sheet, c, r), t.static))
                if not offenders:
                    continue
                key = f"{cell.ref}:{arg.raw}"
                if key in reported:
                    continue
                reported.add(key)
                total = sum(_parse_num(v) for _a, v in offenders)
                yield Finding(
                    rule="R008",
                    title=(f"A number is stored as text inside {arg.raw}" if len(offenders) == 1
                           else f"{len(offenders)} numbers are stored as text inside {arg.raw}"),
                    severity=WARNING,
                    cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                    formula=cell.formula,
                    detail=(f"{', '.join(a for a, _v in offenders[:3])} "
                            f"{'holds' if len(offenders) == 1 else 'hold'} text that looks like a number. "
                            f"{node.name} skips {'it' if len(offenders) == 1 else 'them'}, "
                            f"so about {total:,.0f} is missing from this result."),
                    evidence={"cells": [a for a, _v in offenders[:10]],
                              "values": [v for _a, v in offenders[:10]],
                              "missing_total": total},
                    impact_value=abs(total),
                    impact_currency=is_currency_format(cell.number_format),
                    confidence=0.85,
                )


def _looks_numeric(s: str) -> bool:
    t = s.strip().replace(",", "").replace("¥", "").replace("$", "").replace("%", "")
    if not t or t in ("-", "."):
        return False
    try:
        float(t)
        return True
    except ValueError:
        return False


def _parse_num(s: str) -> float:
    t = s.strip().replace(",", "").replace("¥", "").replace("$", "")
    pct = t.endswith("%")
    t = t.rstrip("%")
    try:
        v = float(t)
    except ValueError:
        return 0.0
    return v / 100 if pct else v


@rule("R010", "Two cells hold the same formula twice",
      "The same calculation in two places drifts apart the moment one is edited.", INFO)
def duplicated_logic(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    from collections import defaultdict

    groups: dict[str, list] = defaultdict(list)
    for cell in wb.formula_cells():
        if cell.ast is None or not cell.formula:
            continue
        if len(cell.formula) < 18:
            continue
        groups[cell.formula.upper()].append(cell)
    for formula, cells in groups.items():
        if len(cells) < 2:
            continue
        distinct_positions = {(c.sheet, c.col) for c in cells} | {(c.sheet, c.row) for c in cells}
        if len(distinct_positions) < 3:
            continue                                  # a copied row or column is normal
        head = cells[0]
        yield Finding(
            rule="R010",
            title=f"Same formula appears in {len(cells)} places",
            severity=INFO,
            cell=head.ref, sheet=head.sheet, col=head.col, row=head.row,
            formula=head.formula,
            detail=(f"{', '.join(c.ref for c in cells[:4])} all contain the identical formula. "
                    "Compute it once and refer to that cell."),
            evidence={"cells": [c.ref for c in cells[:10]]},
            related=[c.ref for c in cells[1:6]],
            confidence=0.6,
        )
