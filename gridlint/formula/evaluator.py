"""Evaluate a parsed formula against a resolver.

Evaluation is a pure function of (node, context sheet, resolver): it never
recurses into other cells' formulas. The engine evaluates cells in dependency
order and feeds already-computed values back through the resolver, which keeps
deep reference chains from ever touching Python's recursion limit.
"""
from __future__ import annotations

from typing import Any, Protocol

from .ast_nodes import (BinOp, ErrorLit, FuncCall, Literal, Name, Node,
                        PostfixOp, RangeRef, Ref, UnaryOp)
from .functions import FUNCTIONS, Matrix
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


def _arith(op: str, a: Any, b: Any) -> Any:
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
            return to_text(_scalar(a)) + to_text(_scalar(b))
        c = compare(_scalar(a), _scalar(b))
        return {"=": c == 0, "<>": c != 0, "<": c < 0, "<=": c <= 0, ">": c > 0, ">=": c >= 0}[op]
    if isinstance(node, FuncCall):
        fn = FUNCTIONS.get(node.name)
        if fn is None:
            raise ExcelError(NAME)
        if node.name in ("IF", "IFERROR", "IFNA", "ISERROR"):
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
    if name == "ISERROR":
        try:
            v = _scalar(_eval(node.args[0], ctx, r))
        except ExcelError:
            return True
        return isinstance(v, ExcelError)
    raise ExcelError(VALUE)
