"""Defects about values being wrong right now: errors, cycles, and stale numbers.

R013 is the rule only a tool with its own calculation engine can run: it
recomputes the workbook and compares against the numbers the spreadsheet app
saved, which catches files that were edited with automatic calculation off.
"""
from __future__ import annotations

from typing import Any, Iterable

from ..formula.ast_nodes import FuncCall, RangeRef, Ref, walk
from ..formula.evaluator import CIRC
from ..formula.functions import VOLATILE
from ..formula.values import BLANK, ExcelError
from ..workbook import Workbook
from .base import CRITICAL, INFO, WARNING, Finding, addr, is_currency_format, rule

ERROR_MEANING = {
    "#DIV/0!": "something is divided by zero or by an empty cell",
    "#VALUE!": "a calculation was given text where it needed a number",
    "#REF!": "a formula points at a cell or sheet that no longer exists",
    "#NAME?": "a function or name in the formula is not recognised",
    "#N/A": "a lookup found nothing",
    "#NUM!": "the result is not a valid number",
    "#NULL!": "two ranges that were expected to overlap do not",
    CIRC: "the formula depends on itself",
}


@rule("R005", "Formula produces an error",
      "An error in one cell spreads to everything computed from it.", CRITICAL)
def error_values(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    errored = {k: v for k, v in computed.items() if isinstance(v, ExcelError)}
    if not errored:
        return
    for key, err in sorted(errored.items()):
        cell = wb.get(*key)
        if cell is None:
            continue
        if err.code == CIRC:
            continue                                  # reported by R006 instead
        root = _root_cause(wb, computed, graph, key)
        is_root = root == key
        downstream = graph.downstream([key])
        yield Finding(
            rule="R005",
            title=f"{err.code} in {cell.addr}" if is_root else f"{err.code} inherited from {addr(*root)}",
            severity=CRITICAL if is_root else WARNING,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=(f"This cell shows {err.code}: {ERROR_MEANING.get(err.code, 'the formula cannot be computed')}."
                    + ("" if is_root else f" The problem starts at {addr(*root)}.")),
            evidence={"error": err.code, "root_cause": addr(*root),
                      "affected_cells": len(downstream)},
            related=[] if is_root else [addr(*root)],
            impact_cells=len(downstream),
            confidence=1.0,
        )


def _root_cause(wb: Workbook, computed: dict, graph, key) -> Any:
    """Walk back through precedents to the first cell that errors on its own."""
    seen = set()
    stack = [key]
    best = key
    while stack:
        k = stack.pop()
        if k in seen:
            continue
        seen.add(k)
        prec = graph.precedents.get(k)
        if not prec:
            continue
        erroring = [p for p in prec.cells if isinstance(computed.get(p), ExcelError)]
        if not erroring:
            best = k
            continue
        for p in erroring:
            stack.append(p)
            best = p
    return best


@rule("R006", "Circular reference",
      "A formula that depends on itself never settles on a trustworthy number.", CRITICAL)
def circular(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cycle in graph.cycles:
        names = [addr(*k) for k in cycle]
        head = wb.get(*cycle[0])
        if head is None:
            continue
        yield Finding(
            rule="R006",
            title="Circular reference",
            severity=CRITICAL,
            cell=head.ref, sheet=head.sheet, col=head.col, row=head.row,
            formula=head.formula,
            detail=(f"{len(cycle)} cells depend on each other in a loop: "
                    f"{' -> '.join(names[:5])}{' -> ...' if len(names) > 5 else ''}."),
            evidence={"cycle": names[:40], "size": len(cycle)},
            related=names[1:8],
            confidence=1.0,
        )


@rule("R013", "Saved value does not match the formula",
      "The file was saved with calculation off, so people are reading numbers the formulas no longer produce.",
      CRITICAL)
def stale_values(wb: Workbook, computed: dict, graph, self_check=None, **_: Any) -> Iterable[Finding]:
    # Only trustworthy when the engine reproduced essentially every other cached value.
    if self_check is None or not self_check.checked:
        return
    mismatches = self_check.mismatches
    if not mismatches:
        return
    agreement_excluding = (self_check.matched) / max(1, self_check.checked - len(mismatches))
    if agreement_excluding < 0.995 or len(mismatches) > max(20, self_check.checked * 0.25):
        return                                        # too many: distrust the engine, not the file
    for m in mismatches:
        sheet_name, _, a1_addr = m["cell"].partition("!")
        cell = next((c for c in wb.formula_cells() if c.ref == m["cell"]), None)
        if cell is None:
            continue
        yield Finding(
            rule="R013",
            title="Displayed number is out of date",
            severity=CRITICAL,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=(f"The file shows {m['expected']}, but this formula now computes {m['got']}. "
                    "The workbook was saved without recalculating."),
            evidence={"saved": m["expected"], "recomputed": m["got"],
                      "engine_agreement": round(self_check.agreement, 4)},
            impact_currency=is_currency_format(cell.number_format),
            confidence=0.9,
        )


@rule("R007", "Formula points at something that is gone",
      "A deleted row, column or sheet leaves #REF! behind, and the number silently changes.", CRITICAL)
def broken_reference(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cell in wb.formula_cells():
        if not cell.formula:
            continue
        if "#REF!" in cell.formula:
            yield Finding(
                rule="R007",
                title="Formula contains a deleted reference",
                severity=CRITICAL,
                cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                formula=cell.formula,
                detail="This formula still refers to a cell, row or sheet that was deleted.",
                evidence={"formula": cell.formula},
                confidence=1.0,
            )
            continue
        if cell.ast is None:
            continue
        for node in walk(cell.ast):
            sheet = getattr(node, "sheet", None)
            if sheet and sheet not in wb.sheets:
                yield Finding(
                    rule="R007",
                    title=f"Formula reads a sheet named {sheet!r} that is not in this file",
                    severity=CRITICAL,
                    cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                    formula=cell.formula,
                    detail=(f"This formula reads from a sheet called {sheet!r}. "
                            "It is not in this workbook, so the value comes from a stale copy or a broken link."),
                    evidence={"missing_sheet": sheet},
                    confidence=0.95,
                )
                break


@rule("R011", "Result changes every time the file opens",
      "Volatile functions make a number nobody can reproduce or audit.", WARNING)
def volatile_functions(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        used = {n.name for n in walk(cell.ast) if isinstance(n, FuncCall) and n.name in VOLATILE}
        if not used:
            continue
        opaque = used & {"INDIRECT", "OFFSET"}
        downstream = len(graph.downstream([(cell.sheet, cell.col, cell.row)]))
        yield Finding(
            rule="R011",
            title=f"{', '.join(sorted(used))} makes this cell unreproducible",
            severity=WARNING,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=("This formula uses " + ", ".join(sorted(used)) +
                    ", so the number can change on its own between two people opening the same file."
                    + (" It also hides which cells it reads, so nothing downstream can be traced."
                       if opaque else "")),
            evidence={"functions": sorted(used), "opaque": sorted(opaque), "downstream": downstream},
            impact_cells=downstream,
            confidence=1.0,
        )


@rule("R012", "IFERROR is hiding a real failure",
      "Wrapping a broken formula in IFERROR replaces an obvious error with a wrong number.", CRITICAL)
def iferror_masking(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    from ..formula.evaluator import evaluate
    from ..engine import _Resolver

    resolver = _Resolver(wb, computed)
    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        for node in walk(cell.ast):
            if not isinstance(node, FuncCall) or node.name not in ("IFERROR", "IFNA"):
                continue
            if not node.args:
                continue
            inner = evaluate(node.args[0], cell.sheet, resolver)
            if not isinstance(inner, ExcelError):
                continue
            fallback = computed.get((cell.sheet, cell.col, cell.row))
            yield Finding(
                rule="R012",
                title=f"{node.name} is masking {inner.code}",
                severity=CRITICAL,
                cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                formula=cell.formula,
                detail=(f"Without the {node.name}, this cell would show {inner.code}: "
                        f"{ERROR_MEANING.get(inner.code, 'the formula cannot be computed')}. "
                        f"Instead it quietly shows {_plain(fallback)}."),
                evidence={"hidden_error": inner.code, "shown_instead": _plain(fallback)},
                confidence=0.95,
            )
            break


def _plain(v: Any) -> Any:
    if isinstance(v, ExcelError):
        return v.code
    if v is BLANK:
        return "(blank)"
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    return v
