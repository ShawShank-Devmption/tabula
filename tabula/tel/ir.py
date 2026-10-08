"""Edit IR and its optimiser (design.md sections 8.4-8.5).

The IR is a flat list of quadruple-like operations: (operator, sheet,
address, operand).  Sheet names are as of that point in program order.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass


@dataclass
class SetCell:
    sheet: str
    col: int
    row: int
    value: object = None      # literal value (when formula is None)
    formula: object = None    # formula AST, already fill-shifted for this cell
    line: int = 0


@dataclass
class ClearCell:
    sheet: str
    col: int
    row: int
    line: int = 0


@dataclass
class Structural:
    verb: str   # 'insert' | 'delete'
    axis: str   # 'rows' | 'cols'
    sheet: str
    at: int
    n: int
    line: int = 0


@dataclass
class AddSheet:
    name: str
    line: int = 0


@dataclass
class RenameSheet:
    old: str
    new: str
    line: int = 0


@dataclass
class Expect:
    host: str | None   # sheet that unqualified references in the condition refer to
    ast: object
    source: str
    line: int = 0
    col: int = 0


def eliminate_dead_writes(ops: list) -> tuple[list, Counter]:
    """Dead-store elimination.

    Scans backwards keeping the set of cells written later.  A write to a cell
    already in that set can never be observed, so it is dropped.  Reads
    (Expect) and address-space changes (structural and sheet ops) are
    barriers that empty the set.  Returns (live ops, Counter{(dead_line,
    killer_line): cells}).
    """
    later: dict[tuple, int] = {}
    live: list = []
    dead: Counter = Counter()
    for op in reversed(ops):
        if isinstance(op, (SetCell, ClearCell)):
            key = (op.sheet.casefold(), op.col, op.row)
            if key in later:
                dead[(op.line, later[key])] += 1
                continue
            later[key] = op.line
        else:
            later.clear()
        live.append(op)
    live.reverse()
    return live, dead
