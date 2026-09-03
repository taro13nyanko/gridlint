"""Evaluate a parsed formula against a resolver.

Evaluation is a pure function of (node, context sheet, resolver): it never
recurses into other cells' formulas. The engine evaluates cells in dependency
order and feeds already-computed values back through the resolver, which keeps
deep reference chains from ever touching Python's recursion limit.
"""
from __future__ import annotations

from typing import Any, Protocol

from .ast_nodes import (BinOp, ErrorLit, ExternalRef, FuncCall, Literal, Name, Node,
                        PostfixOp, RangeRef, Ref, TableRef, UnaryOp)
from .functions import FUNCTIONS, UNSUPPORTED, Matrix
from .values import (BLANK, DIV0, NAME, NUM, REF, VALUE, Blank, ExcelError,
                     compare, to_bool, to_number, to_text)

#: Excel reports circular references separately from ordinary errors.
CIRC = "#CIRC!"
MAX_RANGE_CELLS = 2_000_000


class Resolver(Protocol):
    def cell(self, sheet: str | None, col: int, row: int) -> Any: ...
    def used_bounds(self, sheet: str | None) -> tuple[int, int]: ...
    def sheet_exists(self, sheet: str) -> bool: ...
    def defined_name(self, name: str) -> Any: ...
    def table_range(self, table: str, column: str | None) -> RangeRef | None: ...


def _range_matrix(node: RangeRef, ctx: str | None, resolver: Resolver) -> Matrix:
    sheet = node.sheet or ctx
    if node.sheet is not None and not resolver.sheet_exists(node.sheet):
        raise ExcelError(REF)
    max_col, max_row = resolver.used_bounds(sheet)
    c2 = min(node.col2, max(max_col, node.col1))
    r2 = min(node.row2, max(max_row, node.row1))
    if (c2 - node.col1 + 1) * (r2 - node.row1 + 1) > MAX_RANGE_CELLS:
        raise ExcelError(NUM)
    return Matrix(
        [[resolver.cell(sheet, c, r) for c in range(node.col1, c2 + 1)]
         for r in range(node.row1, r2 + 1)]
    )


def _scalar(v: Any) -> Any:
    """Collapse a 1x1 Matrix; a wider Matrix used as a scalar is an Excel error."""
    if isinstance(v, Matrix):
        flat = list(v.flat())
        if len(flat) == 1:
            return flat[0]
        raise ExcelError(VALUE)
    return v


def _broadcast(op, a: Any, b: Any):
    """Apply a binary operator element by element when either side is a range.

    This is what makes the SUMPRODUCT((range="x")*range) idiom work, which real
    models lean on constantly because it predates SUMIFS.
    """
    if not isinstance(a, Matrix) and not isinstance(b, Matrix):
        return None
    left = a if isinstance(a, Matrix) else None
    right = b if isinstance(b, Matrix) else None
    rows = len(left or right)
    cols = len((left or right)[0]) if rows else 0
    if left is not None and right is not None and (len(left) != len(right)
                                                   or (rows and len(left[0]) != len(right[0]))):
        raise ExcelError(VALUE)
    out = Matrix()
    for r in range(rows):
        row = []
        for c in range(cols):
            av = left[r][c] if left is not None else a
            bv = right[r][c] if right is not None else b
            row.append(op(av, bv))
        out.append(row)
    return out


def _arith(op: str, a: Any, b: Any) -> Any:
    broadcast = _broadcast(lambda x, y: _arith(op, x, y), a, b)
    if broadcast is not None:
        return broadcast
    x, y = to_number(_scalar(a)), to_number(_scalar(b))
    if op == "+":
        return x + y
    if op == "-":
        return x - y
    if op == "*":
        return x * y
    if op == "/":
        if y == 0:
            raise ExcelError(DIV0)
        return x / y
    if op == "^":
        try:
            r = x ** y
        except (OverflowError, ValueError, ZeroDivisionError) as e:
            raise ExcelError(NUM) from e
        if isinstance(r, complex):
            raise ExcelError(NUM)
        return float(r)
    raise ExcelError(VALUE)


def evaluate(node: Node, ctx_sheet: str | None, resolver: Resolver) -> Any:
    """Evaluate `node`. Excel errors are returned as ExcelError values, not raised."""
    try:
        return _eval(node, ctx_sheet, resolver)
    except ExcelError as e:
        return e
    except RecursionError:
        return ExcelError(NUM)


def _eval(node: Node, ctx: str | None, r: Resolver) -> Any:
    if isinstance(node, Literal):
        return node.value
    if isinstance(node, ErrorLit):
        raise ExcelError(node.code)
    if isinstance(node, Ref):
        if node.sheet is not None and not r.sheet_exists(node.sheet):
            raise ExcelError(REF)
        return r.cell(node.sheet or ctx, node.col, node.row)
    if isinstance(node, RangeRef):
        return _range_matrix(node, ctx, r)
    if isinstance(node, TableRef):
        rng = r.table_range(node.table, node.column) if hasattr(r, "table_range") else None
        if rng is None:
            raise ExcelError(NAME)
        return _range_matrix(rng, ctx, r)
    if isinstance(node, ExternalRef):
        # The value lives in a workbook this file does not contain. Guessing it
        # would be worse than admitting we cannot compute it.
        raise ExcelError(NAME)
    if isinstance(node, Name):
        v = r.defined_name(node.name)
        if v is None:
            raise ExcelError(NAME)
        return _eval(v, ctx, r) if isinstance(v, Node) else v
    if isinstance(node, UnaryOp):
        return -to_number(_scalar(_eval(node.operand, ctx, r)))
    if isinstance(node, PostfixOp):
        return to_number(_scalar(_eval(node.operand, ctx, r))) / 100.0
    if isinstance(node, BinOp):
        a = _eval(node.left, ctx, r)
        b = _eval(node.right, ctx, r)
        op = node.op
        if op in ("+", "-", "*", "/", "^"):
            return _arith(op, a, b)
        if op == "&":
            joined = _broadcast(lambda x, y: to_text(x) + to_text(y), a, b)
            return joined if joined is not None else to_text(_scalar(a)) + to_text(_scalar(b))

        def cmp(x, y):
            c = compare(x, y)
            return {"=": c == 0, "<>": c != 0, "<": c < 0,
                    "<=": c <= 0, ">": c > 0, ">=": c >= 0}[op]

        compared = _broadcast(cmp, a, b)
        return compared if compared is not None else cmp(_scalar(a), _scalar(b))
    if isinstance(node, FuncCall):
        fn = FUNCTIONS.get(node.name)
        if fn is None:
            # Both an unknown name and a function Excel has but Gridlint does not
            # model land here. Either way the honest answer is "no value", which
            # the self-check counts as unmodelled rather than as a disagreement.
            raise ExcelError(NAME)
        if node.name in ("IF", "IFERROR", "IFNA", "ISERROR", "ISNA", "ISERR"):
            return _eval_lazy(node, ctx, r)
        args = [_eval(a, ctx, r) for a in node.args]
        args = [a if isinstance(a, Matrix) else _scalar(a) for a in args]
        for a in args:
            if isinstance(a, ExcelError):
                raise a
        return fn(*args)
    raise ExcelError(VALUE)


def _eval_lazy(node: FuncCall, ctx: str | None, r: Resolver) -> Any:
    """IF and the IS*/IFERROR family must not propagate errors from unused branches."""
    name = node.name
    if name == "IF":
        if not node.args:
            raise ExcelError(VALUE)
        cond = to_bool(_scalar(_eval(node.args[0], ctx, r)))
        if cond:
            return _scalar(_eval(node.args[1], ctx, r)) if len(node.args) > 1 else True
        return _scalar(_eval(node.args[2], ctx, r)) if len(node.args) > 2 else False
    if name in ("IFERROR", "IFNA"):
        try:
            v = _scalar(_eval(node.args[0], ctx, r))
        except ExcelError as e:
            v = e
        if isinstance(v, ExcelError) and (name == "IFERROR" or v.code == "#N/A"):
            return _scalar(_eval(node.args[1], ctx, r)) if len(node.args) > 1 else BLANK
        return v
    if name in ("ISERROR", "ISNA", "ISERR"):
        try:
            v = _scalar(_eval(node.args[0], ctx, r))
        except ExcelError as e:
            v = e
        if not isinstance(v, ExcelError):
            return False
        if name == "ISNA":
            return v.code == "#N/A"
        if name == "ISERR":
            return v.code != "#N/A"
        return True
    raise ExcelError(VALUE)
