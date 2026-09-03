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
from .base import CRITICAL, INFO, WARNING, Finding, Fix, addr, is_currency_format, rule

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
        if cell.uses_external or _reads_external(wb, graph, computed, key):
            continue                                  # reported by R014, which says why
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


def _reads_external(wb: Workbook, graph, computed: dict, key) -> bool:
    """True when this cell's error is inherited from a link to another workbook."""
    seen, stack = set(), [key]
    while stack:
        k = stack.pop()
        if k in seen:
            continue
        seen.add(k)
        c = wb.get(*k)
        if c is not None and c.uses_external:
            return True
        for p in graph.precedents.get(k, None).cells if graph.precedents.get(k) else ():
            if isinstance(computed.get(p), ExcelError):
                stack.append(p)
    return False


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


@rule("R014", "The number comes from a workbook you do not have",
      "A link to another file shows whatever that file held the last time somebody opened both.",
      CRITICAL)
def external_links(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    from ..formula.ast_nodes import ExternalRef

    by_target: dict[str, list] = {}
    for cell in wb.formula_cells():
        if cell.ast is None or not cell.uses_external:
            continue
        for node in walk(cell.ast):
            if isinstance(node, ExternalRef):
                book = node.raw.split("]")[0].lstrip("[")
                book = f"another workbook (link {book})" if book.isdigit() else book
                by_target.setdefault(book, []).append((cell, node.raw))

    for book, entries in sorted(by_target.items()):
        cell, raw = entries[0]
        others = len(entries) - 1
        downstream = graph.downstream([(c.sheet, c.col, c.row) for c, _r in entries])
        yield Finding(
            rule="R014",
            title=("Reads a workbook that is not part of this file"
                   if book.startswith("another workbook") else f"Reads a workbook that is not here: {book}"),
            severity=CRITICAL,
            cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
            formula=cell.formula,
            detail=(f"{raw} points into {book}, which is not part of this file. "
                    f"The number shown is a snapshot from the last time somebody had both files "
                    f"open{f', and {others} other formula(s) do the same' if others else ''}. "
                    "Nothing here can check whether it is still right."),
            evidence={"workbook": book, "reference": raw,
                      "formulas": [c.ref for c, _r in entries[:10]], "count": len(entries),
                      "affected_cells": len(downstream)},
            related=[c.ref for c, _r in entries[1:6]],
            impact_cells=len(downstream),
            group_size=len(entries),
            group_cells=[c.ref for c, _r in entries[:24]],
            confidence=1.0,
        )


@rule("R015", "A hidden row or column sits inside a total",
      "A total that spans hidden rows still adds them, and nobody reading the sheet can see why it is high.",
      WARNING)
def hidden_rows_in_range(wb: Workbook, computed: dict, graph, **_: Any) -> Iterable[Finding]:
    from ..formula.ast_nodes import FuncCall, RangeRef
    from ..rules.structural import BLOCK_AGGREGATES

    seen: set[str] = set()
    for cell in wb.formula_cells():
        if cell.ast is None:
            continue
        for node in walk(cell.ast):
            if not isinstance(node, FuncCall) or node.name not in BLOCK_AGGREGATES:
                continue
            for arg in node.args:
                if not isinstance(arg, RangeRef):
                    continue
                sheet_name = arg.sheet or cell.sheet
                sheet = wb.sheets.get(sheet_name)
                if sheet is None or not sheet.hidden_rows:
                    continue
                hidden = sorted(r for r in sheet.hidden_rows if arg.row1 <= r <= arg.row2)
                hidden = [r for r in hidden if wb.get(sheet_name, arg.col1, r) is not None]
                if not hidden:
                    continue
                key = f"{cell.ref}:{arg.raw}"
                if key in seen:
                    continue
                seen.add(key)
                total = sum(_number(_value_at(wb, computed, sheet_name, arg.col1, r)) for r in hidden)
                yield Finding(
                    rule="R015",
                    title=f"{len(hidden)} hidden row(s) inside {arg.raw}",
                    severity=WARNING,
                    cell=cell.ref, sheet=cell.sheet, col=cell.col, row=cell.row,
                    formula=cell.formula,
                    detail=(f"Row{'s' if len(hidden) > 1 else ''} "
                            f"{', '.join(str(r) for r in hidden[:5])} "
                            f"{'are' if len(hidden) > 1 else 'is'} hidden but still counted by this "
                            f"{node.name}, contributing about {total:,.0f}. Someone reading the sheet "
                            "cannot see where that part of the total comes from."),
                    evidence={"range": arg.raw, "hidden_rows": hidden[:20],
                              "hidden_contribution": total},
                    impact_value=abs(total) or None,
                    impact_currency=is_currency_format(cell.number_format),
                    confidence=0.8,
                )


def _value_at(wb: Workbook, computed: dict, sheet: str, col: int, row: int) -> Any:
    key = (sheet, col, row)
    if key in computed:
        return computed[key]
    c = wb.get(sheet, col, row)
    if c is None:
        return None
    return c.static if c.static is not None else c.cached


def _number(v: Any) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def _plain(v: Any) -> Any:
    if isinstance(v, ExcelError):
        return v.code
    if v is BLANK:
        return "(blank)"
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    return v
