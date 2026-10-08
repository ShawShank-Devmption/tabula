"""Simulation and analysis: execute the edit IR on the in-memory workbook
(design.md section 8.6) and report what the edit would do -- before any file
is touched.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .. import parser as P
from ..diagnostics import error, warning
from ..emitter import emit
from ..lexer import LexError
from ..parser import ParseError, parse_target
from ..refs import MAX_COL, MAX_ROW, index_to_col, quote_sheet
from ..relocate import move_position, move_span, relocate, rename_sheet
from ..values import BLANK, UNKNOWN, display, is_error, is_number, values_equal
from ..workbook import Cell, formula_cell, is_simulable, literal_cell
from . import ir


@dataclass
class Change:
    address: str
    before_raw: str
    before_value: object
    after_raw: str
    after_value: object
    unverified: bool = False


@dataclass
class SimResult:
    diagnostics: list = field(default_factory=list)
    writes: list = field(default_factory=list)
    affected: list = field(default_factory=list)
    new_errors: list = field(default_factory=list)
    unverified: list = field(default_factory=list)
    expects: list = field(default_factory=list)
    relocated: int = 0
    structure: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)


class Allowlist:
    """--allow entries: '*', a whole sheet ('Sales'), or a range ('Sales!A1:F40')."""

    def __init__(self, specs):
        self.everything = False
        self.sheets: set[str] = set()
        self.ranges: list[tuple[str, tuple]] = []
        for spec in specs:
            spec = spec.strip()
            if spec == "*":
                self.everything = True
                continue
            if "!" not in spec:
                self.sheets.add(spec.strip("'").casefold())
                continue
            try:
                node = parse_target(spec)
            except (LexError, ParseError) as e:
                raise ValueError(f"bad --allow entry {spec!r}: {e.message}")
            if not isinstance(node, (P.Ref, P.RangeNode)) or node.sheet is None:
                raise ValueError(f"bad --allow entry {spec!r}: use Sheet or Sheet!A1:F40")
            b = ((node.ref.col, node.ref.row, node.ref.col, node.ref.row)
                 if isinstance(node, P.Ref) else node.rng.bounds())
            self.ranges.append((node.sheet.casefold(), b))

    def cell(self, sheet: str, col: int, row: int) -> bool:
        cf = sheet.casefold()
        return self.everything or cf in self.sheets or any(
            s == cf and b[0] <= col <= b[2] and b[1] <= row <= b[3] for s, b in self.ranges)

    def whole_sheet(self, sheet: str) -> bool:
        return self.everything or sheet.casefold() in self.sheets


def address(sheet_name: str, col: int, row: int) -> str:
    return f"{quote_sheet(sheet_name)}!{index_to_col(col)}{row}"


def simulate(engine, ops: list, allow: Allowlist | None = None) -> SimResult:
    wb = engine.wb
    res = SimResult()
    diags = res.diagnostics
    before = {cell.uid: (cell.raw, cell.value) for _, _, _, cell in wb.all_cells()}
    cyclic_before = {engine.cell_at(k).uid for k in engine.cyclic}
    write_line: dict[int, int] = {}
    overwrites: Counter = Counter()
    denied: dict[int, str] = {}
    pending: list[list] = []          # [Expect op, ast (relocated as we go), host Sheet]
    seeds: set = set()
    edited: set = set()               # sheets whose rows/columns moved
    rebuild = False
    unparsed = sum(1 for _, _, _, c in wb.all_cells() if c.kind == "unparsed")
    unparsed += sum(1 for dn in wb.names.values() if dn.ast is None)

    for op in ops:
        if isinstance(op, (ir.SetCell, ir.ClearCell)):
            sheet = wb.sheet(op.sheet)
            if allow is not None and not allow.cell(sheet.name, op.col, op.row):
                denied.setdefault(op.line, address(sheet.name, op.col, op.row))
                continue
            old = sheet.cells.get((op.col, op.row))
            clearing = isinstance(op, ir.ClearCell) or (op.formula is None and op.value == "")
            if clearing:
                if old is None:
                    continue
                new = Cell(raw="", kind="literal", value=BLANK)
            elif op.formula is not None:
                new = formula_cell(op.formula)
            else:
                new = literal_cell(op.value)
                if isinstance(op.value, str):
                    new.raw = op.value
            if old is not None:
                new.uid = old.uid
                if old.kind != "literal" and new.kind == "literal" and not clearing:
                    overwrites[op.line] += 1
            new.touched = True
            sheet.cells[(op.col, op.row)] = new
            write_line[new.uid] = op.line
            if not rebuild:
                engine.register(sheet, op.col, op.row)
                seeds.add((sheet.sid, op.col, op.row))
        elif isinstance(op, ir.Structural):
            sheet = wb.sheet(op.sheet)
            blockers = _address_blockers(wb)
            if blockers or unparsed:
                why = ("the workbook contains " + ", ".join(blockers) if blockers else
                       f"the workbook has {unparsed} formula(s)/name(s) Tabula cannot parse, "
                       "which could not be relocated")
                diags.append(error("E-STRUCT", f"cannot {op.verb} {op.axis} safely: {why}",
                                   op.line, hint="Tabula refuses rather than leave references "
                                                 "pointing at the wrong cells"))
                continue
            if allow is not None and not allow.whole_sheet(sheet.name):
                diags.append(error("E-PERM", f"{op.verb} {op.axis} changes all of sheet "
                                   f"'{sheet.name}', which --allow does not cover", op.line))
                continue
            if op.verb == "insert" and _insertion_overflows(sheet, op.axis, op.at, op.n):
                diags.append(error("E-STRUCT", f"cannot insert {op.axis} safely: stored cells "
                                   "or formatting would move beyond Excel's grid", op.line))
                continue
            n = op.n if op.verb == "insert" else -op.n
            _relocate_all(wb, pending, sheet.name, op.axis, op.at, n)
            _move_cells(sheet, op.axis, op.at, n)
            edited.add(sheet.sid)
            res.structure.append((f"{op.verb}_{op.axis}", sheet.name, op.at, op.n))
            rebuild = True
        elif isinstance(op, ir.AddSheet):
            if allow is not None and not allow.everything:
                diags.append(error("E-PERM", "add sheet needs --allow '*'", op.line))
                continue
            wb.add_sheet(op.name)
            res.structure.append(("add_sheet", op.name))
            rebuild = True
        elif isinstance(op, ir.RenameSheet):
            if allow is not None and not allow.everything:
                diags.append(error("E-PERM", "rename sheet needs --allow '*'", op.line))
                break  # later ops were resolved against the renamed sheet table
            sheet = wb.sheet(op.old)
            blockers = _address_blockers(wb)
            if blockers or unparsed:
                why = ", ".join(blockers) if blockers else (
                    f"the workbook has {unparsed} formula(s)/name(s) Tabula cannot parse")
                diags.append(error("E-STRUCT", f"cannot rename sheet safely: {why}", op.line,
                                   hint="Tabula cannot relocate these stored references"))
                break
            _rename_all(wb, pending, sheet.name, op.new)
            res.structure.append(("rename_sheet", sheet.name, op.new))
            sheet.name = op.new
            rebuild = True
        elif isinstance(op, ir.Expect):
            pending.append([op, op.ast, wb.sheet(op.host) if op.host else None])

    for line, n in sorted(overwrites.items()):
        diags.append(warning("W-OVERWRITE-FORMULA", f"{n} cell(s) that held a formula now hold "
                             "a constant (hard-coded value)", line,
                             hint="write a formula instead if the value should stay live"))
    for line, where in sorted(denied.items()):
        diags.append(error("E-PERM", f"writes outside the allowed ranges (first: {where})", line))

    # -- recompute ---------------------------------------------------------
    if rebuild:
        engine.rebuild()
        seeds = {(s.sid, c, r) for s in wb.sheets for (c, r), cell in s.cells.items()
                 if cell.touched or cell.text_changed}
        # Formulas whose text did not change can still see different data after rows
        # move -- whole-column ranges, relocated defined names -- so recompute every
        # formula that reads an edited sheet.
        seeds |= {key for key, (cells, ranges) in engine.graph.precedents.items()
                  if any(c[0] in edited for c in cells) or any(sid in edited for sid, _ in ranges)}
    stats = engine.recompute(seeds, changed=bool(res.structure or seeds))

    # -- new circular references ---------------------------------------------
    reported: set = set()
    for key in sorted(engine.cyclic):
        cell = engine.cell_at(key)
        if key in reported or cell is None or cell.uid in cyclic_before:
            continue
        path = engine.graph.cycle_path(key) or [key, key]
        reported.update(path)
        lines = [write_line[engine.cell_at(k).uid] for k in path
                 if engine.cell_at(k).uid in write_line]
        diags.append(error("E-CYCLE", "this edit creates a circular reference: "
                           + " -> ".join(engine.address(k) for k in path),
                           min(lines) if lines else 0,
                           hint="a formula must not depend on its own result"))

    # -- what changed ----------------------------------------------------------
    dirty = engine.last_dirty
    for sheet in wb.sheets:
        for (col, row), cell in sorted(sheet.cells.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            addr = address(sheet.name, col, row)
            b_raw, b_val = before.get(cell.uid, ("", BLANK))
            unverified = cell.unverified or cell.value is UNKNOWN
            if cell.touched:
                if cell.raw == "" and b_raw == "":
                    continue
                res.writes.append(Change(addr, b_raw, b_val, cell.raw, cell.value, unverified))
            elif cell.kind in ("formula", "unparsed"):
                if cell.kind == "formula" and cell.text_changed:
                    res.relocated += 1
                if cell.uid in before and not values_equal(cell.value, b_val):
                    res.affected.append(Change(addr, b_raw, b_val, cell.raw, cell.value, unverified))
                elif unverified and (sheet.sid, col, row) in dirty:
                    res.unverified.append(addr)
    for ch in res.writes:
        if ch.unverified and ch.after_raw.startswith("="):
            res.unverified.append(ch.address)
    for ch in res.affected:
        if ch.unverified:
            res.unverified.append(ch.address)
    res.new_errors = [ch for ch in res.writes + res.affected
                      if is_error(ch.after_value) and not is_error(ch.before_value)]
    if res.new_errors:
        sample = ", ".join(f"{c.address} ({display(c.after_value)})" for c in res.new_errors[:5])
        more = f" and {len(res.new_errors) - 5} more" if len(res.new_errors) > 5 else ""
        diags.append(warning("W-NEW-ERROR", f"{len(res.new_errors)} cell(s) now evaluate to "
                             f"an error: {sample}{more}"))

    # -- expectations ----------------------------------------------------------
    for op, ast, host in pending:
        host_name = host.name if host is not None else wb.sheets[0].name
        value, unverified = engine.evaluate_expr(ast, host_name)
        entry = {"line": op.line, "source": op.source, "value": display(value)}
        if unverified:
            entry["status"] = "unverified"
            diags.append(warning("W-EXPECT-UNVERIFIED", "cannot verify this expectation: it "
                                 "depends on values Tabula does not simulate", op.line, op.col))
        elif value is True or (is_number(value) and value != 0):
            entry["status"] = "pass"
        else:
            entry["status"] = "fail"
            detail = _explain_failure(engine, ast, host_name, value)
            diags.append(error("E-EXPECT", f"expectation failed: {op.source} {detail}",
                               op.line, op.col, len(op.source)))
        res.expects.append(entry)

    res.stats = dict(stats, structural=rebuild)
    return res


def _explain_failure(engine, ast, host: str, value) -> str:
    if is_error(value):
        return f"evaluates to {value.code}"
    if isinstance(ast, P.Binary) and ast.op in ("=", "<>", "<", "<=", ">", ">="):
        left, _ = engine.evaluate_expr(ast.left, host)
        right, _ = engine.evaluate_expr(ast.right, host)
        return f"is FALSE (left side {display(left)}, right side {display(right)})"
    if not isinstance(value, bool) and not is_number(value):
        return f"is not a TRUE/FALSE condition (value: {display(value)!r})"
    return "is FALSE"


def _relocate_all(wb, pending, sheet_name: str, axis: str, at: int, n: int) -> None:
    for s in wb.sheets:
        for cell in s.cells.values():
            if cell.kind == "formula":
                new, changed = relocate(cell.ast, s.name, sheet_name, axis, at, n)
                if changed:
                    cell.ast, cell.raw, cell.text_changed = new, emit(new), True
    for dn in wb.names.values():
        if dn.ast is not None:
            host = wb.by_sid(dn.scope).name if dn.scope else None
            new, changed = relocate(dn.ast, host, sheet_name, axis, at, n)
            if changed:
                dn.ast, dn.text, dn.changed = new, emit(new, with_equals=False), True
    for item in pending:
        host = item[2].name if item[2] is not None else None
        item[1], _ = relocate(item[1], host, sheet_name, axis, at, n)


def _move_cells(sheet, axis: str, at: int, n: int) -> None:
    rows = axis == "rows"
    limit = MAX_ROW if rows else MAX_COL
    moved = {}
    for (col, row), cell in sheet.cells.items():
        q = move_position(row if rows else col, at, n, limit)
        if q is not None:
            moved[(col, q) if rows else (q, row)] = cell
    sheet.cells = moved
    moved_grid = set()
    for col, row in getattr(sheet, "grid_cells", ()):
        q = move_position(row if rows else col, at, n, limit)
        if q is not None:
            moved_grid.add((col, q) if rows else (q, row))
    sheet.grid_cells = moved_grid
    attr = "grid_rows" if rows else "grid_cols"
    setattr(sheet, attr, [span for lo, hi in getattr(sheet, attr, ())
                         if (span := move_span(lo, hi, at, n, limit)) is not None])


def _address_blockers(wb):
    # Rules can refer indirectly through names or text expressions. Do not infer
    # absence of dependencies by a regex; conservatively refuse workbook-wide.
    blockers = [f"{what} on sheet '{sheet.name}'" for sheet in wb.sheets
                for what in sheet.structural_blockers()]
    for sheet in wb.sheets:
        count = sum(cell.kind == "formula" and not cell.simulable
                    for cell in sheet.cells.values())
        if count:
            blockers.append(f"{count} unsupported formula(s) on sheet '{sheet.name}' "
                            "whose addresses cannot be verified")
    count = sum(dn.ast is not None and not is_simulable(dn.ast, from_file=True)
                for dn in wb.names.values())
    if count:
        blockers.append(f"{count} unsupported defined name(s) whose addresses cannot be verified")
    return blockers


def _insertion_overflows(sheet, axis, at, n):
    rows = axis == "rows"
    limit = MAX_ROW if rows else MAX_COL
    positions = set(sheet.cells) | set(getattr(sheet, "grid_cells", ()))
    coordinates = (row if rows else col for col, row in positions)
    if any(p >= at and p + n > limit for p in coordinates):
        return True
    return any(hi >= at and hi + n > limit
               for _, hi in getattr(sheet, "grid_rows" if rows else "grid_cols", ()))


def _rename_all(wb, pending, old: str, new: str) -> None:
    for s in wb.sheets:
        for cell in s.cells.values():
            if cell.kind == "formula":
                ast, changed = rename_sheet(cell.ast, old, new)
                if changed:
                    cell.ast, cell.raw, cell.text_changed = ast, emit(ast), True
    for dn in wb.names.values():
        if dn.ast is not None:
            ast, changed = rename_sheet(dn.ast, old, new)
            if changed:
                dn.ast, dn.text, dn.changed = ast, emit(ast, with_equals=False), True
    for item in pending:
        item[1], _ = rename_sheet(item[1], old, new)
