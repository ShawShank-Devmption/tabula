"""Reference relocation over formula ASTs (design.md section 5.5).

Three pure transforms, each returning (new_ast, changed):

  shift_for_fill   Excel fill/copy: relative parts move by an offset, '$' parts stay.
  relocate         insert/delete rows or columns: every address at or past the edit
                   point moves; ranges spanning an insertion expand; ranges that lose
                   cells shrink; references to deleted cells become #REF!.
  rename_sheet     explicit sheet qualifiers follow a renamed sheet.

This is the linker's relocation problem: when the address space changes,
every stored address must be rewritten.
"""
from __future__ import annotations

from dataclasses import replace

from . import parser as P
from .refs import MAX_COL, MAX_ROW, CellRef, RangeRef


class FillError(Exception):
    """A fill pushed a reference off the grid."""


def _transform(node, fn):
    """Rebuild an AST, letting fn replace reference/name leaves. Returns (ast, changed)."""
    if isinstance(node, (P.Ref, P.RangeNode, P.Name)):
        new = fn(node)
        return new, new is not node
    if isinstance(node, P.Unary):
        operand, ch = _transform(node.operand, fn)
        return (P.Unary(node.op, operand), True) if ch else (node, False)
    if isinstance(node, P.Percent):
        operand, ch = _transform(node.operand, fn)
        return (P.Percent(operand), True) if ch else (node, False)
    if isinstance(node, P.Binary):
        left, c1 = _transform(node.left, fn)
        right, c2 = _transform(node.right, fn)
        return (P.Binary(node.op, left, right), True) if (c1 or c2) else (node, False)
    if isinstance(node, P.Call):
        args, changed = [], False
        for a in node.args:
            new, ch = _transform(a, fn)
            args.append(new)
            changed |= ch
        return (P.Call(node.fname, args), True) if changed else (node, False)
    return node, False


# ---------------------------------------------------------------- fill
def _shift_cell(ref: CellRef, dc: int, dr: int) -> CellRef:
    col = ref.col if ref.abs_col else ref.col + dc
    row = ref.row if ref.abs_row else ref.row + dr
    if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
        raise FillError(f"filling moves {ref} off the grid")
    return CellRef(col, row, ref.abs_col, ref.abs_row)


def shift_for_fill(ast, dc: int, dr: int):
    if dc == 0 and dr == 0:
        return ast, False

    def fn(node):
        if isinstance(node, P.Ref):
            return P.Ref(_shift_cell(node.ref, dc, dr), node.sheet)
        if isinstance(node, P.RangeNode):
            rng = node.rng
            if rng.kind == "cols":
                s, e = _shift_cell(rng.start, dc, 0), _shift_cell(rng.end, dc, 0)
            elif rng.kind == "rows":
                s, e = _shift_cell(rng.start, 0, dr), _shift_cell(rng.end, 0, dr)
            else:
                s, e = _shift_cell(rng.start, dc, dr), _shift_cell(rng.end, dc, dr)
            return P.RangeNode(RangeRef(s, e, rng.kind), node.sheet)
        return node

    return _transform(ast, fn)


# ---------------------------------------------------------------- insert / delete
_REF = object()  # marker: position was deleted


def _move_point(p: int, at: int, n: int, limit: int):
    """New coordinate for a single position, or _REF if it was deleted."""
    if n > 0:
        if p < at:
            return p
        return p + n if p + n <= limit else _REF
    k = -n
    end = at + k - 1
    if p < at:
        return p
    if p <= end:
        return _REF
    return p - k


def _move_span(p1: int, p2: int, at: int, n: int, limit: int):
    """New (p1, p2) for a span, or _REF if every position was deleted."""
    if n > 0:
        if p1 >= at:
            return (p1 + n, min(p2 + n, limit)) if p1 + n <= limit else _REF
        if p2 >= at:  # the span contains the insertion point: expand
            return p1, min(p2 + n, limit)
        return p1, p2
    k = -n
    end = at + k - 1
    if p1 >= at and p2 <= end:
        return _REF
    new1 = p1 if p1 < at else (at if p1 <= end else p1 - k)
    new2 = p2 if p2 < at else (at - 1 if p2 <= end else p2 - k)
    return new1, new2


def _same_sheet(node_sheet, host_sheet, sheet) -> bool:
    effective = node_sheet if node_sheet is not None else host_sheet
    return effective is not None and effective.casefold() == sheet.casefold()


def relocate_ref_node(node, host_sheet: str | None, sheet: str, axis: str, at: int, n: int):
    """Relocate one Ref/RangeNode; returns the same object when unaffected."""
    if not isinstance(node, (P.Ref, P.RangeNode)) or not _same_sheet(node.sheet, host_sheet, sheet):
        return node
    rows = axis == "rows"
    limit = MAX_ROW if rows else MAX_COL
    if isinstance(node, P.Ref):
        ref = node.ref
        p = ref.row if rows else ref.col
        q = _move_point(p, at, n, limit)
        if q is _REF:
            return P.ErrorLit("#REF!")
        if q == p:
            return node
        new = replace(ref, row=q) if rows else replace(ref, col=q)
        return P.Ref(new, node.sheet)
    rng = node.rng
    if (rows and rng.kind == "cols") or (not rows and rng.kind == "rows"):
        return node  # full extent along the edited axis: unaffected
    s, e = rng.start, rng.end
    p1, p2 = (s.row, e.row) if rows else (s.col, e.col)
    swapped = p1 > p2
    lo, hi = (p2, p1) if swapped else (p1, p2)
    moved = _move_span(lo, hi, at, n, limit)
    if moved is _REF:
        return P.ErrorLit("#REF!")
    if moved == (lo, hi):
        return node
    new_lo, new_hi = moved
    if swapped:
        new_lo, new_hi = new_hi, new_lo
    if rows:
        s2, e2 = replace(s, row=new_lo), replace(e, row=new_hi)
    else:
        s2, e2 = replace(s, col=new_lo), replace(e, col=new_hi)
    return P.RangeNode(RangeRef(s2, e2, rng.kind), node.sheet)


def relocate(ast, host_sheet: str | None, sheet: str, axis: str, at: int, n: int):
    """Insert (n > 0) or delete (n < 0, starting at `at`) rows/cols on `sheet`."""
    return _transform(ast, lambda node: relocate_ref_node(node, host_sheet, sheet, axis, at, n))


# ---------------------------------------------------------------- rename
def rename_sheet(ast, old: str, new: str):
    def fn(node):
        if node.sheet is not None and node.sheet.casefold() == old.casefold():
            return replace(node, sheet=new)
        return node

    return _transform(ast, fn)


def move_position(p: int, at: int, n: int, limit: int) -> int | None:
    """Where a row/column index ends up after an insert/delete; None if deleted."""
    q = _move_point(p, at, n, limit)
    return None if q is _REF else q


def move_span(p1: int, p2: int, at: int, n: int, limit: int) -> tuple[int, int] | None:
    """Where a span of rows/columns ends up after an insert/delete; None if deleted."""
    moved = _move_span(p1, p2, at, n, limit)
    return None if moved is _REF else moved
