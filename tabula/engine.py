"""Workbook engine: dependency-driven recomputation (design.md section 6.3).

Engine    recompute(seeds) = closure -> Kahn order (Tarjan for cycles) -> evaluate,
          with unverified-value (taint) propagation for cells Tabula cannot compute.
Sheet     the Phase 1 REPL facade: one in-memory sheet with set/get/show/explain.
"""
from __future__ import annotations

import math
import time

from .evaluator import Evaluator
from .graph import DependencyGraph
from .lexer import LexError, tokenize
from .parser import ParseError, dump_ast, parse_formula
from .refs import CellRef, index_to_col
from .values import (BLANK, CYCLE_ERR, NUM_ERR, REF_ERR, UNKNOWN, VALUE_ERR, display,
                     is_number, values_equal)
from .workbook import MISSING, Cell, Workbook, formula_cell, is_simulable, literal_cell, parse_literal


class Engine:
    def __init__(self, workbook: Workbook):
        self.wb = workbook
        self.graph = DependencyGraph()
        self.evaluator = Evaluator(self)
        self.cyclic: set = set()
        self.last_stats: dict = {}
        self.last_dirty: set = set()
        self._saw_unknown = False
        self._saw_unverified = False

    # ------------------------------------------------------------ evaluator context
    def _resolve_sheet(self, sheet: str | None, host: str):
        return self.wb.sheet(sheet if sheet is not None else host)

    def _read(self, cell: Cell | None):
        if cell is None:
            return BLANK
        if cell.value is UNKNOWN:
            self._saw_unknown = True
        if cell.unverified:
            self._saw_unverified = True
        return cell.value

    def value(self, sheet, host, col, row):
        s = self._resolve_sheet(sheet, host)
        if s is None:
            return REF_ERR
        return self._read(s.cells.get((col, row)))

    def range_cells(self, sheet, host, rng):
        s = self._resolve_sheet(sheet, host)
        if s is None:
            return REF_ERR
        c1, r1, c2, r2 = rng.bounds()
        if rng.size() <= len(s.cells):
            found = ((c, r, s.cells.get((c, r))) for r in range(r1, r2 + 1)
                     for c in range(c1, c2 + 1))
            found = [(c, r, cell) for c, r, cell in found if cell is not None]
        else:
            found = sorted(((c, r, cell) for (c, r), cell in s.cells.items()
                            if c1 <= c <= c2 and r1 <= r <= r2), key=lambda t: (t[1], t[0]))
        out = []
        for c, r, cell in found:
            v = self._read(cell)
            if v is not BLANK:
                out.append((c, r, v))
        return out

    def resolve_name(self, node, host):
        h = self.wb.sheet(host)
        dn = self.wb.lookup_name(node.ident, h.sid if h else None, node.sheet)
        return dn.ast if dn is not None else None

    # ------------------------------------------------------------ graph maintenance
    def register(self, sheet, col: int, row: int) -> None:
        key = (sheet.sid, col, row)
        cell = sheet.cells.get((col, row))
        if cell is not None and cell.kind == "formula":
            cell.simulable = cell.simulable and self.wb.names_simulable(sheet, cell.ast)
            cells, ranges = self.wb.precedents(sheet, cell.ast)
            self.graph.set_formula(key, cells, ranges)
        else:
            self.graph.remove_formula(key)

    def rebuild(self) -> None:
        self.graph = DependencyGraph()
        self.cyclic = set()  # keys may have moved; cycles are re-found by the next recompute
        for sheet, col, row, _ in self.wb.formulas():
            self.register(sheet, col, row)

    def cell_at(self, key) -> Cell | None:
        sid, col, row = key
        return self.wb.by_sid(sid).cells.get((col, row))

    # ------------------------------------------------------------ recomputation
    def recompute(self, seeds=None, initial: bool = False, changed: bool = False) -> dict:
        """Recompute `seeds` and their transitive dependents (all formulas if None).

        initial=True is the first computation after loading: unsimulable formulas
        keep Excel's cached value and count as verified.  Afterwards, an
        unsimulable formula whose inputs may have changed is marked unverified.
        """
        start = time.perf_counter()
        seeds = None if seeds is None else set(seeds)
        invalidated = set()
        if not initial and (changed or seeds is None or seeds):
            # Unknown function/name semantics can hide reads (INDIRECT, OFFSET,
            # spill/array syntax). Explicit AST edges are not proof of complete
            # dependencies, so any edit invalidates these cached results.
            for sheet, col, row, cell in self.wb.all_cells():
                if cell.kind == "unparsed" or (cell.kind == "formula" and not cell.simulable):
                    cell.unverified = True
                    invalidated.add((sheet.sid, col, row))
        dirty = (set(self.graph.precedents) if seeds is None else
                 self.graph.closure(seeds | invalidated))
        self.last_dirty = dirty | invalidated
        order, cyclic = self.graph.order(dirty)
        self.cyclic = (self.cyclic - dirty) | cyclic
        for key in cyclic:
            cell = self.cell_at(key)
            cell.value, cell.unverified = CYCLE_ERR, False
        for key in order:
            cell = self.cell_at(key)
            sheet = self.wb.by_sid(key[0])
            if not cell.simulable:
                if initial:
                    cell.value = UNKNOWN if cell.cached is MISSING else cell.cached
                    cell.unverified = cell.cached is MISSING
                else:
                    cell.unverified = True
                continue
            self._saw_unknown = self._saw_unverified = False
            v = self.evaluator.evaluate(cell.ast, sheet.name)
            if v is BLANK:
                v = 0.0  # a formula that reads an empty cell shows 0, as in Excel
            elif is_number(v) and not math.isfinite(v):
                v = NUM_ERR
            if self._saw_unknown:
                v = UNKNOWN
            cell.value = v
            cell.unverified = self._saw_unknown or self._saw_unverified
        self.last_stats = {
            "formulas_recomputed": len(dirty),
            "total_formulas": len(self.graph.precedents),
            "recompute_ms": round((time.perf_counter() - start) * 1000, 3),
        }
        return self.last_stats

    def evaluate_expr(self, ast, host: str):
        """(value, unverified) of an ad-hoc expression, e.g. a TEL `expect`."""
        sheet = self.wb.sheet(host)
        if not is_simulable(ast, True) or not self.wb.names_simulable(sheet, ast):
            return UNKNOWN, True
        self._saw_unknown = self._saw_unverified = False
        v = self.evaluator.evaluate(ast, host)
        if self._saw_unknown:
            v = UNKNOWN
        return v, self._saw_unknown or self._saw_unverified

    def agreement(self) -> tuple[int, int, list]:
        """(matching, compared, mismatches) of Tabula values vs Excel-cached values."""
        matching = compared = 0
        mismatches = []
        for sheet, col, row, cell in self.wb.formulas():
            if not cell.simulable or cell.cached is MISSING or cell.unverified:
                continue
            compared += 1
            if values_equal(cell.value, cell.cached):
                matching += 1
            else:
                mismatches.append((sheet.name, col, row, cell.value, cell.cached))
        return matching, compared, mismatches

    # ------------------------------------------------------------ presentation
    def address(self, key) -> str:
        from .refs import quote_sheet
        sid, col, row = key
        return f"{quote_sheet(self.wb.by_sid(sid).name)}!{index_to_col(col)}{row}"

    def cycle_text(self, key) -> str:
        path = self.graph.cycle_path(key) or [key, key]
        return " -> ".join(self.address(k) for k in path)


# ---------------------------------------------------------------- REPL facade
def caret_diagnostic(cellname: str, category: str, message: str,
                     formula: str, pos: int) -> str:
    # pos is within the formula body; +1 accounts for the leading '='
    return (
        f"error[{category}] in {cellname}: {message}\n"
        f"  ={formula}\n"
        f"  {' ' * (pos + 1)}^"
    )


class Sheet:
    """A single in-memory sheet for the interactive REPL and its tests."""

    def __init__(self):
        self.wb = Workbook()
        self.sheet = self.wb.add_sheet("Sheet1")
        self.engine = Engine(self.wb)

    @property
    def cells(self) -> dict:
        return self.sheet.cells

    def set(self, refname: str, raw: str) -> str | None:
        """Set a cell. Returns a diagnostic string on compile error, else None."""
        ref = CellRef.parse(refname)
        if ref is None:
            return f"error[RefError]: invalid cell reference '{refname}'"
        raw = raw.strip()
        diag = None
        if raw.startswith("="):
            body = raw[1:]
            try:
                cell = formula_cell(parse_formula(body), raw)
            except (LexError, ParseError) as e:
                category = "LexicalError" if isinstance(e, LexError) else "SyntaxError"
                diag = caret_diagnostic(str(ref), category, e.message, body, e.pos)
                cell = Cell(raw=raw, kind="unparsed", value=VALUE_ERR)  # contained: shows #VALUE!
        else:
            cell = literal_cell(parse_literal(raw))
            cell.raw = raw
        self.sheet.cells[ref.key()] = cell
        self.engine.register(self.sheet, ref.col, ref.row)
        self.engine.recompute({(self.sheet.sid, ref.col, ref.row)})
        return diag

    def value(self, refname: str):
        ref = CellRef.parse(refname)
        cell = self.sheet.cells.get(ref.key()) if ref else None
        return BLANK if cell is None else cell.value

    def show(self) -> str:
        if not self.sheet.cells:
            return "(empty sheet)"
        cols = sorted({c for c, _ in self.sheet.cells})
        rows = sorted({r for _, r in self.sheet.cells})
        table = {(c, r): display(self.value(f"{index_to_col(c)}{r}")) for c in cols for r in rows}
        widths = {c: max([len(index_to_col(c))] + [len(table[(c, r)]) for r in rows]) + 2
                  for c in cols}
        lines = ["     " + "".join(index_to_col(c).ljust(widths[c]) for c in cols)]
        for r in rows:
            lines.append((f"{r:<5}" + "".join(table[(c, r)].ljust(widths[c]) for c in cols)).rstrip())
        return "\n".join(lines)

    def explain(self, refname: str) -> str:
        ref = CellRef.parse(refname)
        cell = self.sheet.cells.get(ref.key()) if ref else None
        if cell is None:
            return f"{refname}: (blank)"
        return "\n".join(explain_cell(str(ref), cell, self.engine, self.sheet))


def explain_cell(name: str, cell: Cell, engine: Engine, sheet) -> list[str]:
    out = [f"cell   : {name}", f"raw    : {cell.raw}", f"kind   : {cell.kind}"]
    if cell.kind == "formula":
        toks = [t for t in tokenize(cell.raw) if t.lexeme]
        out.append("tokens : " + " ".join(t.lexeme for t in toks))
        out.append("types  : " + " ".join(t.type.name for t in toks))
        out.append("ast    :")
        out.extend("  " + line for line in dump_ast(cell.ast))
        cells, ranges = engine.wb.precedents(sheet, cell.ast)
        refs = sorted(engine.address(k) for k in cells)
        refs += sorted(f"{engine.address((sid, b[0], b[1]))}:{index_to_col(b[2])}{b[3]}"
                       for sid, b in ranges)
        out.append("reads  : " + (", ".join(refs) if refs else "(nothing)"))
        if not cell.simulable:
            out.append("note   : uses a function Tabula does not simulate; value from Excel's cache")
    out.append(f"value  : {display(cell.value)}" + (" (unverified)" if cell.unverified else ""))
    if cell.cached is not MISSING:
        out.append(f"excel  : {display(cell.cached)}  (cached in the file)")
    return out
