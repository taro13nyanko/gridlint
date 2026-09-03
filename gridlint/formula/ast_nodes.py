"""AST node types for parsed formulas."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class Node:
    pass


@dataclass(frozen=True)
class Literal(Node):
    value: Any          # float | str | bool
    raw: str = ""


@dataclass(frozen=True)
class ErrorLit(Node):
    code: str


@dataclass(frozen=True)
class Ref(Node):
    """A single cell reference. `sheet` is None for same-sheet refs."""
    sheet: str | None
    col: int            # 1-based
    row: int            # 1-based
    abs_col: bool = False
    abs_row: bool = False
    raw: str = ""


@dataclass(frozen=True)
class RangeRef(Node):
    sheet: str | None
    col1: int
    row1: int
    col2: int
    row2: int
    raw: str = ""

    @property
    def n_cells(self) -> int:
        return (self.col2 - self.col1 + 1) * (self.row2 - self.row1 + 1)


@dataclass(frozen=True)
class Name(Node):
    name: str


@dataclass(frozen=True)
class UnaryOp(Node):
    op: str
    operand: Node


@dataclass(frozen=True)
class PostfixOp(Node):
    op: str             # only "%"
    operand: Node


@dataclass(frozen=True)
class BinOp(Node):
    op: str
    left: Node
    right: Node


@dataclass(frozen=True)
class FuncCall(Node):
    name: str
    args: tuple[Node, ...] = field(default_factory=tuple)


def walk(node: Node):
    """Yield every node in the tree, parents before children."""
    yield node
    for f in ("operand", "left", "right"):
        child = getattr(node, f, None)
        if isinstance(child, Node):
            yield from walk(child)
    for arg in getattr(node, "args", ()) or ():
        if isinstance(arg, Node):
            yield from walk(arg)
