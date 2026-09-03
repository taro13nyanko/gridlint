"""Precedence-climbing parser for Excel formulas.

Precedence follows Excel, including its two surprises:
  * unary minus binds TIGHTER than "^", so -2^2 = 4 (not -4)
  * "^" is LEFT-associative, so 2^3^2 = 64 (not 512)
"""
from __future__ import annotations

import re

from .ast_nodes import BinOp, ErrorLit, FuncCall, Literal, Name, Node, PostfixOp, RangeRef, Ref, UnaryOp
from .tokenizer import FormulaSyntaxError, Token, TokType, tokenize

# binary operator precedence (higher binds tighter)
_PREC = {
    "=": 1, "<>": 1, "<": 1, "<=": 1, ">": 1, ">=": 1,
    "&": 2,
    "+": 3, "-": 3,
    "*": 4, "/": 4,
    "^": 5,
}
_MAX_COL = 16384
_MAX_ROW = 1048576


def col_to_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(index: int) -> str:
    out = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


_REF_PARTS = re.compile(r"^(?:(.*)!)?(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")


def parse_ref(raw: str) -> Ref:
    m = _REF_PARTS.match(raw)
    if not m:
        raise FormulaSyntaxError(f"bad reference {raw!r}")
    sheet, ac, col, ar, row = m.groups()
    ci, ri = col_to_index(col), int(row)
    if ci > _MAX_COL or ri > _MAX_ROW:
        raise FormulaSyntaxError(f"reference out of bounds: {raw!r}")
    return Ref(sheet=sheet, col=ci, row=ri, abs_col=ac == "$", abs_row=ar == "$", raw=raw)


def parse_range(raw: str) -> RangeRef:
    sheet = None
    body = raw
    if "!" in raw:
        sheet, body = raw.rsplit("!", 1)
    a, _, b = body.partition(":")
    a_s, b_s = a.replace("$", ""), b.replace("$", "")
    if a_s.isdigit() and b_s.isdigit():          # whole-row range 3:7
        r1, r2 = int(a_s), int(b_s)
        return RangeRef(sheet, 1, min(r1, r2), _MAX_COL, max(r1, r2), raw)
    if a_s.isalpha() and b_s.isalpha():          # whole-column range B:D
        c1, c2 = col_to_index(a_s), col_to_index(b_s)
        return RangeRef(sheet, min(c1, c2), 1, max(c1, c2), _MAX_ROW, raw)
    r1 = parse_ref(f"{sheet}!{a}" if sheet else a)
    r2 = parse_ref(f"{sheet}!{b}" if sheet else b)
    return RangeRef(sheet, min(r1.col, r2.col), min(r1.row, r2.row),
                    max(r1.col, r2.col), max(r1.row, r2.row), raw)


class Parser:
    def __init__(self, tokens: list[Token]):
        self.toks = tokens
        self.i = 0

    def peek(self) -> Token | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> Token:
        t = self.peek()
        if t is None:
            raise FormulaSyntaxError("unexpected end of formula")
        self.i += 1
        return t

    def expect(self, ttype: TokType) -> Token:
        t = self.next()
        if t.type is not ttype:
            raise FormulaSyntaxError(f"expected {ttype.value} but found {t.value!r} at {t.pos}")
        return t

    def parse(self) -> Node:
        node = self.parse_expr(0)
        if self.peek() is not None:
            t = self.peek()
            raise FormulaSyntaxError(f"trailing input {t.value!r} at {t.pos}")
        return node

    def parse_expr(self, min_prec: int) -> Node:
        left = self.parse_unary()
        while True:
            t = self.peek()
            if t is None or t.type is not TokType.OP or t.value not in _PREC:
                break
            prec = _PREC[t.value]
            if prec < min_prec:
                break
            self.next()
            right = self.parse_expr(prec + 1)   # left-assoc for every operator, "^" included
            left = BinOp(t.value, left, right)
        return left

    def parse_unary(self) -> Node:
        t = self.peek()
        if t is not None and t.type is TokType.OP and t.value in ("-", "+"):
            self.next()
            operand = self.parse_unary()
            node: Node = operand if t.value == "+" else UnaryOp("-", operand)
            return self.parse_postfix(node)
        return self.parse_postfix(self.parse_atom())

    def parse_postfix(self, node: Node) -> Node:
        while True:
            t = self.peek()
            if t is not None and t.type is TokType.OP and t.value == "%":
                self.next()
                node = PostfixOp("%", node)
                continue
            return node

    def parse_atom(self) -> Node:
        t = self.next()
        if t.type is TokType.NUMBER:
            return Literal(float(t.value), t.value)
        if t.type is TokType.STRING:
            return Literal(t.value, t.value)
        if t.type is TokType.BOOL:
            return Literal(t.value == "TRUE", t.value)
        if t.type is TokType.ERROR:
            return ErrorLit(t.value)
        if t.type is TokType.REF:
            return parse_ref(t.value)
        if t.type is TokType.RANGE:
            return parse_range(t.value)
        if t.type is TokType.NAME:
            return Name(t.value)
        if t.type is TokType.LPAREN:
            inner = self.parse_expr(0)
            self.expect(TokType.RPAREN)
            return inner
        if t.type is TokType.FUNC:
            self.expect(TokType.LPAREN)
            args: list[Node] = []
            if self.peek() is not None and self.peek().type is TokType.RPAREN:
                self.next()
                return FuncCall(t.value, ())
            while True:
                nxt = self.peek()
                if nxt is not None and nxt.type in (TokType.COMMA, TokType.SEMI):
                    args.append(Literal("", ""))       # empty argument, e.g. IF(a,,b)
                else:
                    args.append(self.parse_expr(0))
                nxt = self.next()
                if nxt.type is TokType.RPAREN:
                    break
                if nxt.type not in (TokType.COMMA, TokType.SEMI):
                    raise FormulaSyntaxError(f"expected , or ) at {nxt.pos}")
            return FuncCall(t.name if hasattr(t, "name") else t.value, tuple(args))
        raise FormulaSyntaxError(f"unexpected token {t.value!r} at {t.pos}")


def parse_formula(formula: str) -> Node:
    return Parser(tokenize(formula)).parse()
