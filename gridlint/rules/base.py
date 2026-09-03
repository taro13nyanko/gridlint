"""Finding model, rule registry, and the R1C1 normaliser rules share."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Iterable

from ..formula.ast_nodes import (BinOp, ErrorLit, FuncCall, Literal, Name, Node,
                                 PostfixOp, RangeRef, Ref, UnaryOp)
from ..formula.parser import index_to_col
from ..workbook import Workbook

Severity = str          # "critical" | "warning" | "info"

CRITICAL = "critical"
WARNING = "warning"
INFO = "info"

_SEV_ORDER = {CRITICAL: 0, WARNING: 1, INFO: 2}


@dataclass
class Edit:
    cell: str                       # "Sheet!A1"
    new_formula: str
    old_formula: str | None = None


@dataclass
class Fix:
    """A concrete, machine-applicable change: one or more formula replacements."""
    edits: list[Edit] = field(default_factory=list)
    label: str = ""
    kind: str = "formula"

    @classmethod
    def one(cls, cell: str, new_formula: str, label: str = "", old_formula: str | None = None) -> "Fix":
        return cls(edits=[Edit(cell, new_formula, old_formula)], label=label)

    @property
    def cell(self) -> str:
        return self.edits[0].cell if self.edits else ""

    @property
    def new_formula(self) -> str:
        return self.edits[0].new_formula if self.edits else ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "label": self.label, "kind": self.kind,
            "cell": self.cell, "new_formula": self.new_formula,
            "edits": [{"cell": e.cell, "new_formula": e.new_formula, "old_formula": e.old_formula}
                      for e in self.edits],
        }


@dataclass
class Finding:
    rule: str                       # "R001"
    title: str
    severity: Severity
    cell: str                       # "Sheet!C16"
    sheet: str
    col: int
    row: int
    detail: str                     # one plain sentence, no jargon
    formula: str | None = None
    evidence: dict[str, Any] = field(default_factory=dict)
    fix: Fix | None = None
    related: list[str] = field(default_factory=list)
    confidence: float = 1.0
    # filled in by the impact pass
    impact_cells: int = 0
    impact_value: float | None = None
    impact_cell: str | None = None
    impact_before: Any = None
    impact_after: Any = None
    impact_currency: bool = False
    explanation: str | None = None      # written by the language model, optional
    fix_verified: bool | None = None
    fix_diff: list[dict[str, Any]] = field(default_factory=list)
    group_size: int = 1
    group_cells: list[str] = field(default_factory=list)
    headline_changes: list[dict[str, Any]] = field(default_factory=list)
    trace: list[dict[str, Any]] = field(default_factory=list)

    @property
    def id(self) -> str:
        return f"{self.rule}:{self.cell}"

    @property
    def sort_key(self) -> tuple:
        return (_SEV_ORDER.get(self.severity, 3), -(self.impact_value or 0.0),
                -self.impact_cells, -self.confidence, self.cell)

    def to_dict(self) -> dict[str, Any]:
        d = {
            "id": self.id,
            "rule": self.rule, "title": self.title, "severity": self.severity,
            "cell": self.cell, "sheet": self.sheet, "col": self.col, "row": self.row,
            "detail": self.detail, "formula": self.formula, "evidence": self.evidence,
            "related": self.related, "confidence": round(self.confidence, 2),
            "impact_cells": self.impact_cells, "impact_value": self.impact_value,
            "impact_cell": self.impact_cell, "impact_before": self.impact_before,
            "impact_after": self.impact_after, "impact_currency": self.impact_currency,
            "explanation": self.explanation, "fix_verified": self.fix_verified,
            "fix_diff": self.fix_diff[:12],
            "group_size": self.group_size, "group_cells": self.group_cells[:24],
            "headline_changes": self.headline_changes,
            "trace": self.trace[:24],
        }
        if self.fix:
            d["fix"] = self.fix.to_dict()
        return d


@dataclass
class RuleMeta:
    code: str
    name: str
    why: str
    default_severity: Severity


RuleFn = Callable[..., Iterable[Finding]]
_REGISTRY: list[tuple[RuleMeta, RuleFn]] = []


def rule(code: str, name: str, why: str, severity: Severity = WARNING):
    def deco(fn: RuleFn) -> RuleFn:
        _REGISTRY.append((RuleMeta(code, name, why, severity), fn))
        return fn
    return deco


def registry() -> list[tuple[RuleMeta, RuleFn]]:
    return sorted(_REGISTRY, key=lambda kv: kv[0].code)


# ---------------------------------------------------------------------------
# R1C1 normalisation: two formulas that do "the same thing one row down" get the
# same signature, which is how structural outliers are found without guessing.
# ---------------------------------------------------------------------------

def to_r1c1(node: Node, base_col: int, base_row: int) -> str:
    if isinstance(node, Literal):
        if isinstance(node.value, bool):
            return "TRUE" if node.value else "FALSE"
        if isinstance(node.value, str):
            return f'"{node.value}"'
        return _num(node.value)
    if isinstance(node, ErrorLit):
        return node.code
    if isinstance(node, Ref):
        return _ref_r1c1(node, base_col, base_row)
    if isinstance(node, RangeRef):
        a = _ref_r1c1(Ref(node.sheet, node.col1, node.row1, "$" in node.raw, "$" in node.raw), base_col, base_row)
        b = _ref_r1c1(Ref(None, node.col2, node.row2, "$" in node.raw, "$" in node.raw), base_col, base_row)
        return f"{a}:{b}"
    if isinstance(node, Name):
        return node.name.upper()
    if isinstance(node, UnaryOp):
        return f"({node.op}{to_r1c1(node.operand, base_col, base_row)})"
    if isinstance(node, PostfixOp):
        return f"({to_r1c1(node.operand, base_col, base_row)}{node.op})"
    if isinstance(node, BinOp):
        return f"({to_r1c1(node.left, base_col, base_row)}{node.op}{to_r1c1(node.right, base_col, base_row)})"
    if isinstance(node, FuncCall):
        inner = ",".join(to_r1c1(a, base_col, base_row) for a in node.args)
        return f"{node.name}({inner})"
    return "?"


def _num(v: float) -> str:
    return str(int(v)) if float(v) == int(v) else repr(float(v))


def _ref_r1c1(ref: Ref, base_col: int, base_row: int) -> str:
    sheet = f"{ref.sheet}!" if ref.sheet else ""
    r = f"R{ref.row}" if ref.abs_row else f"R[{ref.row - base_row}]"
    c = f"C{ref.col}" if ref.abs_col else f"C[{ref.col - base_col}]"
    return f"{sheet}{r}{c}"


def shape_of(node: Node | None, base_col: int, base_row: int) -> str:
    """Signature ignoring literal values, so =B5*1.08 and =B6*1.10 look alike."""
    if node is None:
        return ""
    return _shape(node, base_col, base_row)


def _shape(node: Node, bc: int, br: int) -> str:
    if isinstance(node, Literal):
        return "#" if isinstance(node.value, (int, float)) and not isinstance(node.value, bool) else "$str"
    if isinstance(node, (Ref, RangeRef)):
        return to_r1c1(node, bc, br)
    if isinstance(node, UnaryOp):
        return f"({node.op}{_shape(node.operand, bc, br)})"
    if isinstance(node, PostfixOp):
        return f"({_shape(node.operand, bc, br)}{node.op})"
    if isinstance(node, BinOp):
        return f"({_shape(node.left, bc, br)}{node.op}{_shape(node.right, bc, br)})"
    if isinstance(node, FuncCall):
        return f"{node.name}({','.join(_shape(a, bc, br) for a in node.args)})"
    if isinstance(node, ErrorLit):
        return node.code
    if isinstance(node, Name):
        return node.name.upper()
    return "?"


CURRENCY_HINTS = ("¥", "$", "€", "£", "円", "USD", "JPY", "EUR", "GBP", "#,##")


def is_currency_format(fmt: str) -> bool:
    f = fmt or ""
    return any(h in f for h in CURRENCY_HINTS)


def format_kind(fmt: str) -> str:
    """How a value should be shown back to the user: as money, a percentage, or plain.

    A gross margin of 0.063 means nothing on screen; 6.3% is the number the
    person actually recognises from their own sheet.
    """
    f = fmt or ""
    if "%" in f:
        return "percent"
    if is_currency_format(f):
        return "currency"
    return "plain"


def addr(sheet: str, col: int, row: int) -> str:
    return f"{sheet}!{index_to_col(col)}{row}"


def skeleton_of(node: Node | None) -> str:
    """Signature with every reference and literal erased: two cells that do the
    same *kind* of calculation match even when they point at different cells."""
    if node is None:
        return ""
    return _skeleton(node)


def _skeleton(node: Node) -> str:
    if isinstance(node, (Ref, RangeRef)):
        return "@"
    if isinstance(node, Literal):
        return "#" if isinstance(node.value, (int, float)) and not isinstance(node.value, bool) else "$str"
    if isinstance(node, UnaryOp):
        return f"({node.op}{_skeleton(node.operand)})"
    if isinstance(node, PostfixOp):
        return f"({_skeleton(node.operand)}{node.op})"
    if isinstance(node, BinOp):
        return f"({_skeleton(node.left)}{node.op}{_skeleton(node.right)})"
    if isinstance(node, FuncCall):
        return f"{node.name}({','.join(_skeleton(a) for a in node.args)})"
    if isinstance(node, ErrorLit):
        return node.code
    if isinstance(node, Name):
        return node.name.upper()
    return "?"
