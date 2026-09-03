"""Tokenizer for Excel formulas.

Deliberately small and total: every formula either tokenizes or raises
`FormulaSyntaxError`. No regex backtracking traps, single left-to-right pass.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum


class TokType(str, Enum):
    NUMBER = "number"
    STRING = "string"
    BOOL = "bool"
    ERROR = "error"
    REF = "ref"          # A1, $A$1, Sheet1!A1, 'My Sheet'!A1
    RANGE = "range"      # A1:B9, Sheet1!A1:B9
    NAME = "name"        # defined name or function name (followed by "(")
    FUNC = "func"        # NAME immediately followed by "("
    OP = "op"
    LPAREN = "("
    RPAREN = ")"
    COMMA = ","
    SEMI = ";"


class FormulaSyntaxError(ValueError):
    pass


@dataclass(frozen=True)
class Token:
    type: TokType
    value: str
    pos: int


_ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#N/A", "#SPILL!", "#CALC!")

# order matters: longest operators first
_OPS = ("<=", ">=", "<>", "+", "-", "*", "/", "^", "&", "=", "<", ">", "%")

_NUM_RE = re.compile(r"\d+(\.\d*)?([eE][+-]?\d+)?|\.\d+([eE][+-]?\d+)?")
# sheet prefix: Name! or 'Name with spaces'!  (also Sheet1:Sheet3! 3-D refs -> unsupported, rejected)
_SHEET_RE = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!")
_A1_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})(?![A-Za-z0-9_])")
_COL_RANGE_RE = re.compile(r"\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![A-Za-z0-9_])")
_ROW_RANGE_RE = re.compile(r"\$?([1-9]\d{0,6}):\$?([1-9]\d{0,6})(?![A-Za-z0-9_])")
_NAME_RE = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")


def tokenize(formula: str) -> list[Token]:
    s = formula.strip()
    if s.startswith("="):
        s = s[1:]
    out: list[Token] = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch in " \t\r\n":
            i += 1
            continue
        if ch == '"':
            j, buf = i + 1, []
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(s[j])
                j += 1
            if j >= n:
                raise FormulaSyntaxError(f"unterminated string at {i}")
            out.append(Token(TokType.STRING, "".join(buf), i))
            i = j + 1
            continue
        if ch == "#":
            for e in _ERRORS:
                if s.startswith(e, i):
                    out.append(Token(TokType.ERROR, e, i))
                    i += len(e)
                    break
            else:
                raise FormulaSyntaxError(f"unknown error literal at {i}")
            continue
        if ch == "(":
            out.append(Token(TokType.LPAREN, "(", i)); i += 1; continue
        if ch == ")":
            out.append(Token(TokType.RPAREN, ")", i)); i += 1; continue
        if ch == ",":
            out.append(Token(TokType.COMMA, ",", i)); i += 1; continue
        if ch == ";":
            out.append(Token(TokType.SEMI, ";", i)); i += 1; continue
        if ch == "{":
            raise FormulaSyntaxError("array literals are not supported")

        # sheet-qualified reference
        m_sheet = _SHEET_RE.match(s, i)
        prefix, after = "", i
        if m_sheet:
            sheet = (m_sheet.group(1) or "").replace("''", "'") or m_sheet.group(2)
            prefix = sheet + "!"
            after = m_sheet.end()

        m1 = _A1_RE.match(s, after)
        if m1:
            end = m1.end()
            if end < n and s[end] == ":":
                m2 = _A1_RE.match(s, end + 1)
                if m2:
                    out.append(Token(TokType.RANGE, prefix + s[after:m2.end()], i))
                    i = m2.end()
                    continue
            out.append(Token(TokType.REF, prefix + m1.group(0), i))
            i = end
            continue
        mcr = _COL_RANGE_RE.match(s, after)
        if mcr:
            out.append(Token(TokType.RANGE, prefix + mcr.group(0), i)); i = mcr.end(); continue
        mrr = _ROW_RANGE_RE.match(s, after)
        if mrr:
            out.append(Token(TokType.RANGE, prefix + mrr.group(0), i)); i = mrr.end(); continue
        if m_sheet:
            # a sheet prefix must be followed by a reference
            raise FormulaSyntaxError(f"sheet prefix without reference at {i}")

        mn = _NUM_RE.match(s, i)
        if mn and ch.isdigit() or (ch == "." and mn):
            out.append(Token(TokType.NUMBER, mn.group(0), i)); i = mn.end(); continue

        mname = _NAME_RE.match(s, i)
        if mname:
            word = mname.group(0)
            up = word.upper()
            j = mname.end()
            k = j
            while k < n and s[k] in " \t":
                k += 1
            if k < n and s[k] == "(":
                out.append(Token(TokType.FUNC, up, i))
                i = j
                continue
            if up in ("TRUE", "FALSE"):
                out.append(Token(TokType.BOOL, up, i)); i = j; continue
            out.append(Token(TokType.NAME, word, i)); i = j; continue

        for op in _OPS:
            if s.startswith(op, i):
                out.append(Token(TokType.OP, op, i)); i += len(op); break
        else:
            raise FormulaSyntaxError(f"unexpected character {ch!r} at {i}")
    return out
