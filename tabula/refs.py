"""Cell and range references: A1 parsing/formatting, Excel bounds, sheet names.

Bounds are Excel's: columns A..XFD (16,384) and rows 1..1,048,576.
A RangeRef has a kind: 'cells' (A1:B4), 'cols' (A:C) or 'rows' (1:3);
whole-column and whole-row ranges are stored with their full extent so
containment checks need no special cases.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

MAX_ROW = 1_048_576
MAX_COL = 16_384  # XFD

_REF_RE = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)([0-9]+)$")
_COL_RE = re.compile(r"^(\$?)([A-Z]{1,3})$")
_ROW_RE = re.compile(r"^(\$?)([0-9]+)$")
_PLAIN_SHEET_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_R1C1_RE = re.compile(r"^(?=.)([Rr][0-9]*)?([Cc][0-9]*)?$")  # R, C, RC, R1C1, ...


def col_to_index(letters: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27, 'XFD' -> 16384."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def index_to_col(n: int) -> str:
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def parse_col(text: str) -> tuple[int, bool] | None:
    """'$C' -> (3, True); None if not a column within bounds."""
    m = _COL_RE.match(text.upper())
    if not m:
        return None
    col = col_to_index(m.group(2))
    return (col, m.group(1) == "$") if 1 <= col <= MAX_COL else None


def parse_row(text: str) -> tuple[int, bool] | None:
    """'$7' -> (7, True); None if not a row within bounds."""
    m = _ROW_RE.match(text)
    if not m:
        return None
    row = int(m.group(2))
    return (row, m.group(1) == "$") if 1 <= row <= MAX_ROW else None


def quote_sheet(name: str) -> str:
    """Sheet name as it must appear before '!' in a formula."""
    if (_PLAIN_SHEET_RE.match(name) and not _REF_RE.match(name.upper())
            and not _R1C1_RE.match(name)):
        return name
    return "'" + name.replace("'", "''") + "'"


@dataclass(frozen=True)
class CellRef:
    col: int  # 1-based
    row: int  # 1-based
    abs_col: bool = False
    abs_row: bool = False

    @staticmethod
    def parse(text: str) -> "CellRef | None":
        m = _REF_RE.match(text.upper())
        if not m:
            return None
        dc, letters, dr, digits = m.groups()
        col, row = col_to_index(letters), int(digits)
        if not (1 <= row <= MAX_ROW and 1 <= col <= MAX_COL):
            return None
        return CellRef(col, row, abs_col=dc == "$", abs_row=dr == "$")

    def key(self) -> tuple[int, int]:
        return (self.col, self.row)

    def plain(self) -> "CellRef":
        """The same position without '$' markers."""
        return CellRef(self.col, self.row)

    def __str__(self) -> str:
        return (
            ("$" if self.abs_col else "") + index_to_col(self.col)
            + ("$" if self.abs_row else "") + str(self.row)
        )


@dataclass(frozen=True)
class RangeRef:
    start: CellRef
    end: CellRef
    kind: str = "cells"  # 'cells' | 'cols' | 'rows'

    @staticmethod
    def columns(c1: int, c2: int, abs1: bool = False, abs2: bool = False) -> "RangeRef":
        return RangeRef(CellRef(c1, 1, abs1, False), CellRef(c2, MAX_ROW, abs2, False), "cols")

    @staticmethod
    def rows(r1: int, r2: int, abs1: bool = False, abs2: bool = False) -> "RangeRef":
        return RangeRef(CellRef(1, r1, False, abs1), CellRef(MAX_COL, r2, False, abs2), "rows")

    def bounds(self) -> tuple[int, int, int, int]:
        """(col1, row1, col2, row2), normalised so col1<=col2 and row1<=row2."""
        c1, c2 = sorted((self.start.col, self.end.col))
        r1, r2 = sorted((self.start.row, self.end.row))
        return c1, r1, c2, r2

    def contains(self, col: int, row: int) -> bool:
        c1, r1, c2, r2 = self.bounds()
        return c1 <= col <= c2 and r1 <= row <= r2

    def size(self) -> int:
        c1, r1, c2, r2 = self.bounds()
        return (c2 - c1 + 1) * (r2 - r1 + 1)

    def shape(self) -> tuple[int, int]:
        """(rows, cols)."""
        c1, r1, c2, r2 = self.bounds()
        return (r2 - r1 + 1, c2 - c1 + 1)

    def cells(self):
        c1, r1, c2, r2 = self.bounds()
        for row in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                yield CellRef(col, row)

    def __str__(self) -> str:
        if self.kind == "cols":
            return (("$" if self.start.abs_col else "") + index_to_col(self.start.col) + ":"
                    + ("$" if self.end.abs_col else "") + index_to_col(self.end.col))
        if self.kind == "rows":
            return (("$" if self.start.abs_row else "") + str(self.start.row) + ":"
                    + ("$" if self.end.abs_row else "") + str(self.end.row))
        return f"{self.start}:{self.end}"
