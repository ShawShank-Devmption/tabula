"""Workbook model: sheets, cells and defined names (design.md section 6.1).

The workbook is also the compiler's global symbol table: sheet names and
defined names are resolved here, sheet-scoped names before workbook-scoped
ones (Excel's own two-level scoping).
"""
from __future__ import annotations

import itertools
from dataclasses import dataclass, field

from . import parser as P
from .functions import EXCEL_FUNCTIONS, SUPPORTED
from .values import BLANK, UNKNOWN, is_error, to_number

_uids = itertools.count(1)
MISSING = object()  # no cached value available


@dataclass(eq=False)
class Cell:
    raw: str                      # literal text, or '=...' formula text
    kind: str                     # 'literal' | 'formula' | 'unparsed'
    ast: object = None
    value: object = BLANK
    cached: object = MISSING      # Excel's cached value (formulas loaded from a file)
    simulable: bool = True        # every function call is SUPPORTED with valid arity
    unverified: bool = False      # value may be stale (taint)
    touched: bool = False         # written by a TEL script
    text_changed: bool = False    # formula rewritten by relocation; must be re-emitted
    uid: int = field(default_factory=lambda: next(_uids))

    @property
    def is_formula(self) -> bool:
        return self.kind == "formula"


def literal_cell(value) -> Cell:
    from .values import display
    return Cell(raw=display(value), kind="literal", value=value)


def formula_cell(ast, raw: str | None = None, from_file: bool = False) -> Cell:
    from .emitter import emit
    return Cell(raw=raw if raw is not None else emit(ast), kind="formula", ast=ast,
                value=UNKNOWN, simulable=is_simulable(ast, from_file))


def is_simulable(ast, from_file: bool = False) -> bool:
    """Can the evaluator compute this formula exactly as Excel would?

    A call to a function Excel has but Tabula does not implement blocks
    simulation.  A name unknown to Excel evaluates to #NAME? -- which is
    Excel's answer too -- except in loaded files, where the catalogue might
    simply be missing a newer function, so the cached value is trusted instead.
    """
    for node in P.walk(ast):
        if isinstance(node, P.Call):
            spec = SUPPORTED.get(node.fname)
            if spec is not None:
                if len(node.args) < spec.min_args or (
                        spec.max_args is not None and len(node.args) > spec.max_args):
                    return False
            elif from_file or node.fname in EXCEL_FUNCTIONS:
                return False
    return True


def parse_literal(raw: str):
    """REPL literal: blank, TRUE/FALSE, number, else text."""
    if raw == "":
        return BLANK
    if raw.upper() in ("TRUE", "FALSE"):
        return raw.upper() == "TRUE"
    n = to_number(raw)
    return raw if is_error(n) else n


@dataclass(eq=False)
class Sheet:
    sid: int
    name: str
    cells: dict = field(default_factory=dict)   # (col, row) -> Cell
    meta: dict = field(default_factory=dict)    # counts of features Tabula does not relocate
    rule_texts: list = field(default_factory=list)  # (kind, formula text) of DV / CF rules

    def structural_blockers(self) -> list[str]:
        return [f"{n} {what}" for what, n in sorted(self.meta.items()) if n]

    def rules_referring_to(self, sheet_name: str) -> list[str]:
        """Kinds of this sheet's validation / formatting rules that mention another sheet."""
        import re
        pattern = re.compile(r"(?<![A-Za-z0-9_.])(" + re.escape(sheet_name) + r"|'"
                             + re.escape(sheet_name.replace("'", "''")) + r"')!", re.IGNORECASE)
        return sorted({kind for kind, text in self.rule_texts if pattern.search(text or "")})


@dataclass(eq=False)
class DefinedName:
    name: str                 # as written
    scope: int | None         # sheet id for sheet-scoped names, None = workbook
    ast: object | None        # parsed reference/expression, None if unparsable
    text: str                 # raw text from the file
    changed: bool = False     # rewritten by relocation; must be written back


class Workbook:
    def __init__(self):
        self.sheets: list[Sheet] = []
        self.names: dict[tuple, DefinedName] = {}
        self._next_sid = itertools.count(1)

    # -- sheets ------------------------------------------------------------
    def add_sheet(self, name: str) -> Sheet:
        sheet = Sheet(next(self._next_sid), name)
        self.sheets.append(sheet)
        return sheet

    def sheet(self, name: str | None) -> Sheet | None:
        if name is None:
            return None
        folded = name.casefold()
        for s in self.sheets:
            if s.name.casefold() == folded:
                return s
        return None

    def by_sid(self, sid: int) -> Sheet:
        for s in self.sheets:
            if s.sid == sid:
                return s
        raise KeyError(sid)

    def sheet_index(self, sid: int) -> int:
        for i, s in enumerate(self.sheets):
            if s.sid == sid:
                return i
        return len(self.sheets)

    # -- defined names -----------------------------------------------------
    def define(self, name: str, ast, text: str, scope: int | None = None) -> None:
        self.names[(scope, name.upper())] = DefinedName(name, scope, ast, text)

    def lookup_name(self, ident: str, host_sid: int | None,
                    explicit_sheet: str | None = None) -> DefinedName | None:
        key = ident.upper()
        if explicit_sheet is not None:
            sheet = self.sheet(explicit_sheet)
            return self.names.get((sheet.sid, key)) if sheet else None
        if host_sid is not None and (host_sid, key) in self.names:
            return self.names[(host_sid, key)]
        return self.names.get((None, key))

    # -- iteration ---------------------------------------------------------
    def formulas(self):
        for sheet in self.sheets:
            for (col, row), cell in sheet.cells.items():
                if cell.kind == "formula":
                    yield sheet, col, row, cell

    def all_cells(self):
        for sheet in self.sheets:
            for (col, row), cell in sheet.cells.items():
                yield sheet, col, row, cell

    # -- precedents (used to build the dependency graph) --------------------
    def precedents(self, host: Sheet, ast) -> tuple[set, list]:
        """(cell keys, [(sid, bounds)]) that a formula hosted on `host` reads."""
        cells: set = set()
        ranges: list = []
        pending = [(ast, host)]
        seen_names: set = set()
        while pending:
            node_root, h = pending.pop()
            for node in P.walk(node_root):
                if isinstance(node, P.Ref):
                    sheet = self.sheet(node.sheet) if node.sheet else h
                    if sheet is not None:
                        cells.add((sheet.sid, node.ref.col, node.ref.row))
                elif isinstance(node, P.RangeNode):
                    sheet = self.sheet(node.sheet) if node.sheet else h
                    if sheet is not None:
                        ranges.append((sheet.sid, node.rng.bounds()))
                elif isinstance(node, P.Name):
                    dn = self.lookup_name(node.ident, h.sid if h else None, node.sheet)
                    if dn is not None and dn.ast is not None and id(dn) not in seen_names:
                        seen_names.add(id(dn))
                        scope_sheet = self.by_sid(dn.scope) if dn.scope else h
                        pending.append((dn.ast, scope_sheet))
        return cells, ranges
