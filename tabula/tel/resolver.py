"""TEL semantic analysis and lowering (design.md sections 8.3-8.4).

The resolver walks the TEL AST with a scope chain.  The global scope is the
workbook (sheets and defined names); each `in` block pushes a scope with a
current sheet; `let` declares into the innermost scope.  Every target and
every reference inside a formula is resolved against these symbol tables
before anything executes -- this is where hallucinated sheets, unknown
names, wrong shapes and range-in-scalar-context mistakes are caught.  Valid
statements are lowered to the per-cell edit IR.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace

from .. import parser as P
from ..diagnostics import error, suggest, warning
from ..functions import EXCEL_FUNCTIONS, SUPPORTED, classify
from ..refs import MAX_COL, MAX_ROW, CellRef, index_to_col, quote_sheet
from ..relocate import FillError, relocate_ref_node, shift_for_fill
from . import ir
from . import parser as T

MAX_TARGET_CELLS = 100_000
_LOOKS_LIKE_REF = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]+$")
_BAD_SHEET_CHARS = set("[]:*?/\\")


@dataclass
class Symbol:
    name: str
    kind: str              # 'range' | 'value'
    line: int
    node: object = None    # Ref | RangeNode with .sheet set (kind 'range')
    value: object = None   # literal (kind 'value')
    sheet_cf: str = ""


@dataclass
class Scope:
    sheet: str | None
    symbols: dict = field(default_factory=dict)
    poisoned: bool = False  # the block's sheet did not resolve; suppress cascades


def _fmt_bounds(sheet: str, b) -> str:
    c1, r1, c2, r2 = b
    first = f"{index_to_col(c1)}{r1}"
    tail = "" if (c1, r1) == (c2, r2) else f":{index_to_col(c2)}{r2}"
    return f"{quote_sheet(sheet)}!{first}{tail}"


class Resolver:
    def __init__(self, workbook):
        self.wb = workbook
        self.sheets = {s.name.casefold(): s.name for s in workbook.sheets}
        self.origin = {s.name.casefold(): s.name for s in workbook.sheets}  # current -> file name
        self.scopes = [Scope(None)]
        self.ops: list = []
        self.diags: list = []
        self.statements = 0
        self.struct_line: dict[str, int] = {}
        self.struct_ops: list[tuple] = []  # (file sheet name, axis, at, signed n), in order

    # ------------------------------------------------------------ driver
    def run(self, statements) -> tuple[list, list]:
        self._block(statements)
        return self.ops, self.diags

    def _block(self, statements) -> None:
        for st in statements:
            self.statements += 1
            handler = getattr(self, "_" + type(st).__name__.lower())
            handler(st)

    def _err(self, *a, **k):
        self.diags.append(error(*a, **k))

    def _warn(self, *a, **k):
        self.diags.append(warning(*a, **k))

    # ------------------------------------------------------------ scopes and sheets
    def _lookup(self, name: str) -> Symbol | None:
        key = name.upper()
        for scope in reversed(self.scopes):
            if key in scope.symbols:
                return scope.symbols[key]
        return None

    def _scope_sheet(self) -> tuple[str | None, bool]:
        for scope in reversed(self.scopes):
            if scope.poisoned:
                return None, True
            if scope.sheet is not None:
                return scope.sheet, False
        return None, False

    def _host_sid(self, sheet: str | None):
        if sheet is None:
            return None
        original = self.origin.get(sheet.casefold())
        s = self.wb.sheet(original) if original else None
        return s.sid if s else None

    def _sheet_of(self, explicit: str | None, line: int, col: int, what: str):
        """Display name of the sheet a reference means, or None (diagnostic emitted)."""
        if explicit is not None:
            name = self.sheets.get(explicit.casefold())
            if name is None:
                hint = suggest(explicit, self.sheets.values())
                self._err("E-SHEET", f"unknown sheet '{explicit}'", line, col, len(explicit),
                          f"did you mean '{hint}'?" if hint else
                          "sheets: " + ", ".join(self.sheets.values()))
            return name
        sheet, poisoned = self._scope_sheet()
        if poisoned:
            return None
        if sheet is not None:
            return sheet
        if len(self.sheets) == 1:
            return next(iter(self.sheets.values()))
        first = next(iter(self.sheets.values()))
        self._err("E-NOSHEET", f"{what} has no sheet and the workbook has {len(self.sheets)} sheets",
                  line, col, hint=f"qualify it ({quote_sheet(first)}!A1) or wrap the statement in "
                                  f'in "{first}" {{ ... }}')
        return None

    def _stale_check(self, sym: Symbol, line: int, col: int) -> None:
        if sym.kind == "range" and self.struct_line.get(sym.sheet_cf, 0) > sym.line:
            self._warn("W-STALE-LET", f"'{sym.name}' was bound on line {sym.line}, before rows/"
                       f"columns of its sheet changed on line {self.struct_line[sym.sheet_cf]}; "
                       "it still means the address as written", line, col, len(sym.name),
                       "declare it again after the structural edit if it should follow the data")

    # ------------------------------------------------------------ targets
    def _resolve_target(self, tsrc: T.TargetSrc, line: int):
        node = tsrc.node
        if isinstance(node, P.Name):
            sym = self._lookup(node.ident) if node.sheet is None else None
            if sym is not None:
                if sym.kind == "value":
                    self._err("E-TYPE", f"'{sym.name}' is bound to a value, not a cell or range",
                              line, tsrc.col, len(tsrc.text))
                    return None
                self._stale_check(sym, line, tsrc.col)
                node = sym.node
            else:
                scope_sheet, _ = self._scope_sheet()
                dn = self.wb.lookup_name(node.ident, self._host_sid(scope_sheet),
                                         self.origin.get((node.sheet or "").casefold(), node.sheet))
                if dn is not None and isinstance(dn.ast, (P.Ref, P.RangeNode)) and dn.ast.sheet:
                    current = next((cur for cur, orig in self.origin.items()
                                    if orig and orig.casefold() == dn.ast.sheet.casefold()), None)
                    node = dn.ast
                    # the name moves with its cells: replay earlier structural edits on it
                    for origin_name, axis, at, n in self.struct_ops:
                        if origin_name.casefold() == node.sheet.casefold():
                            node = relocate_ref_node(node, None, node.sheet, axis, at, n)
                            if isinstance(node, P.ErrorLit):
                                self._err("E-REF", f"name '{dn.name}' refers to cells deleted "
                                          "earlier in this script", line, tsrc.col, len(tsrc.text))
                                return None
                    node = replace(node, sheet=self.sheets.get(current, node.sheet))
                else:
                    self._unknown_name(node.ident, line, tsrc.col)
                    return None
        sheet = self._sheet_of(node.sheet, line, tsrc.col, f"target '{tsrc.text}'")
        if sheet is None:
            return None
        if isinstance(node, P.Ref):
            b = (node.ref.col, node.ref.row, node.ref.col, node.ref.row)
        else:
            b = node.rng.bounds()
        size = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
        if size > MAX_TARGET_CELLS:
            self._err("E-SIZE", f"target '{tsrc.text}' covers {size:,} cells "
                      f"(limit {MAX_TARGET_CELLS:,})", line, tsrc.col, len(tsrc.text),
                      "use an explicit range such as A2:A500 instead of whole rows/columns")
            return None
        return sheet, b

    def _unknown_name(self, ident: str, line: int, col: int) -> None:
        if _LOOKS_LIKE_REF.match(ident):
            self._err("E-REF", f"'{ident}' is not a valid cell reference", line, col, len(ident),
                      "rows run 1..1048576 and columns A..XFD")
            return
        candidates = [s.name for scope in self.scopes for s in scope.symbols.values()]
        candidates += [dn.name for dn in self.wb.names.values()]
        hint = suggest(ident, candidates)
        self._err("E-NAME", f"unknown name '{ident}'", line, col, len(ident),
                  f"did you mean '{hint}'?" if hint else
                  "text values need double quotes; names must be declared with let "
                  "or defined in the workbook")

    # ------------------------------------------------------------ formulas
    def _bind_formula(self, fsrc: T.FormulaSrc, host: str | None, line: int):
        """Resolve names/sheets/functions in a formula; inline let symbols."""
        errors_before = sum(d.is_error for d in self.diags)
        scope_sheet, _ = self._scope_sheet()
        state = {"scope_warned": False, "needs_host": False}
        unsimulated: set[str] = set()

        def col_of(needle: str) -> int:
            idx = fsrc.text.upper().find(needle.upper())
            return fsrc.col + idx if idx >= 0 else fsrc.col

        def check_sheet(node):
            if node.sheet is not None:
                if self.sheets.get(node.sheet.casefold()) is None:
                    self._sheet_of(node.sheet, line, col_of(node.sheet), "reference")
                return
            if host is None:
                state["needs_host"] = True
            elif (scope_sheet is not None and host.casefold() != scope_sheet.casefold()
                  and not state["scope_warned"]):
                state["scope_warned"] = True
                self._warn("W-SCOPE", f"this formula is written to sheet '{host}', so its "
                           f"unqualified references mean {quote_sheet(host)}!..., not "
                           f"{quote_sheet(scope_sheet)}!...", line, fsrc.col,
                           hint=f"qualify them, e.g. {quote_sheet(scope_sheet)}!A1, if that "
                                "is what you meant")

        def range_misuse(text: str) -> None:
            self._err("E-TYPE", f"range '{text}' is used where a single value is needed",
                      line, col_of(text.split("!")[-1]), len(text),
                      "aggregate it (SUM, AVERAGE, ...) or set the target range to a formula "
                      "for its first cell -- TEL fills it across the range")

        def visit(node, range_ok: bool):
            if isinstance(node, P.Name):
                sym = self._lookup(node.ident) if node.sheet is None else None
                if sym is not None:
                    self._stale_check(sym, line, col_of(node.ident))
                    if sym.kind == "value":
                        v = sym.value
                        return P.Bool(v) if isinstance(v, bool) else (
                            P.Number(v) if isinstance(v, float) else P.Text(v))
                    inlined = sym.node
                    if host is not None and inlined.sheet.casefold() == host.casefold():
                        inlined = replace(inlined, sheet=None)
                    if isinstance(inlined, P.RangeNode) and not range_ok:
                        range_misuse(sym.name)
                    return inlined
                dn = self.wb.lookup_name(node.ident, self._host_sid(host),
                                         self.origin.get((node.sheet or "").casefold(), node.sheet))
                if dn is None:
                    self._unknown_name(node.ident, line, col_of(node.ident))
                elif isinstance(dn.ast, P.RangeNode) and not range_ok:
                    range_misuse(node.ident)
                return node
            if isinstance(node, P.Ref):
                check_sheet(node)
                return node
            if isinstance(node, P.RangeNode):
                check_sheet(node)
                if not range_ok:
                    range_misuse(_text(node))
                return node
            if isinstance(node, P.Unary):
                return P.Unary(node.op, visit(node.operand, False))
            if isinstance(node, P.Percent):
                return P.Percent(visit(node.operand, False))
            if isinstance(node, P.Binary):
                return P.Binary(node.op, visit(node.left, False), visit(node.right, False))
            if isinstance(node, P.Call):
                kind = classify(node.fname)
                if kind == "unknown":
                    hint = suggest(node.fname, EXCEL_FUNCTIONS)
                    self._err("E-FUNC", f"'{node.fname}' is not an Excel function", line,
                              col_of(node.fname + "("), len(node.fname),
                              f"did you mean {hint}?" if hint else None)
                    return node
                if kind == "unsimulated":
                    unsimulated.add(node.fname)
                    return P.Call(node.fname, [visit(a, True) for a in node.args])
                spec = SUPPORTED[node.fname]
                n = len(node.args)
                if n < spec.min_args or (spec.max_args is not None and n > spec.max_args):
                    want = (f"exactly {spec.min_args}" if spec.min_args == spec.max_args else
                            f"at least {spec.min_args}" if spec.max_args is None else
                            f"{spec.min_args} to {spec.max_args}")
                    self._err("E-ARITY", f"{node.fname} takes {want} argument(s), got {n}",
                              line, col_of(node.fname + "("), len(node.fname))
                return P.Call(node.fname, [visit(a, spec.accepts_range(i))
                                           for i, a in enumerate(node.args)])
            return node

        bound = visit(fsrc.ast, False)
        if state["needs_host"]:
            self._sheet_of(None, line, fsrc.col, "this condition's references")
        for fname in sorted(unsimulated):
            self._warn("W-UNSIMULATED", f"Tabula does not simulate {fname}; cells depending on "
                       "it are reported as unverified (Excel computes them on open)",
                       line, col_of(fname), len(fname))
        if sum(d.is_error for d in self.diags) > errors_before:
            return None
        return bound

    # ------------------------------------------------------------ statements
    def _inblock(self, st: T.InBlock) -> None:
        name = self.sheets.get(st.sheet.casefold())
        if name is None:
            self._sheet_of(st.sheet, st.line, st.col, "block")
            self.scopes.append(Scope(None, poisoned=True))
        else:
            self.scopes.append(Scope(name))
        self._block(st.body)
        self.scopes.pop()

    def _let(self, st: T.Let) -> None:
        key = st.name.upper()
        if CellRef.parse(st.name) is not None or key in ("TRUE", "FALSE") or \
                st.name.lower() in T.KEYWORDS:
            self._err("E-DECL", f"'{st.name}' cannot be used as a name", st.line, st.col,
                      len(st.name), "names must not look like cell references or keywords")
            return
        scope = self.scopes[-1]
        if key in scope.symbols:
            self._err("E-DECL", f"'{st.name}' is already declared on line "
                      f"{scope.symbols[key].line}", st.line, st.col, len(st.name))
            return
        if self._lookup(st.name) is not None:
            self._warn("W-SHADOW", f"'{st.name}' shadows a name from an enclosing block",
                       st.line, st.col, len(st.name))
        if isinstance(st.value, T.Literal):
            scope.symbols[key] = Symbol(st.name, "value", st.line, value=st.value.value)
            return
        resolved = self._resolve_target(st.value, st.line)
        if resolved is None:
            return
        sheet, b = resolved
        c1, r1, c2, r2 = b
        # a let name denotes a fixed address (like a defined name): absolute, so a
        # formula filled across a range keeps pointing at it
        start, end = CellRef(c1, r1, True, True), CellRef(c2, r2, True, True)
        node = (P.Ref(start, sheet) if (c1, r1) == (c2, r2) else
                P.RangeNode(P.RangeRef(start, end), sheet))
        scope.symbols[key] = Symbol(st.name, "range", st.line, node=node,
                                    sheet_cf=sheet.casefold())

    def _set(self, st: T.Set) -> None:
        resolved = self._resolve_target(st.target, st.line)
        if resolved is None:
            # still check the formula, so all its mistakes are reported in one run
            if isinstance(st.value, T.FormulaSrc):
                self._bind_formula(st.value, self._scope_sheet()[0], st.line)
            return
        sheet, (c1, r1, c2, r2) = resolved
        rows, cols = r2 - r1 + 1, c2 - c1 + 1
        value = st.value
        if isinstance(value, T.Literal):
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self.ops.append(ir.SetCell(sheet, c, r, value=value.value, line=st.line))
            return
        if isinstance(value, T.ListLit):
            grid = self._shape(value, rows, cols, st)
            if grid is None:
                return
            for i, r in enumerate(range(r1, r2 + 1)):
                for j, c in enumerate(range(c1, c2 + 1)):
                    self.ops.append(ir.SetCell(sheet, c, r, value=grid[i][j], line=st.line))
            return
        ast = self._bind_formula(value, sheet, st.line)
        if ast is None:
            return
        cells = []
        try:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    shifted, _ = shift_for_fill(ast, c - c1, r - r1)
                    cells.append(ir.SetCell(sheet, c, r, formula=shifted, line=st.line))
        except FillError as e:
            self._err("E-REF", f"filling this formula across {st.target.text}: {e}",
                      st.line, value.col, hint="anchor the reference with $ if it should not move")
            return
        self.ops.extend(cells)

    def _shape(self, value: T.ListLit, rows: int, cols: int, st):
        items = value.rows
        target = f"{st.target.text} ({rows}x{cols})"
        if value.one_dimensional:
            n = len(items[0])
            if (rows, cols) == (1, n):
                return items
            if (rows, cols) == (n, 1):
                return [[v] for v in items[0]]
            self._err("E-SHAPE", f"list has {n} items but target {target} needs "
                      f"{rows * cols}", st.line, value.col,
                      hint="a 1-D list fills one row or one column; use [[...],[...]] for blocks")
            return None
        widths = {len(r) for r in items}
        if len(widths) != 1:
            self._err("E-SHAPE", "rows of the list have different lengths", st.line, value.col)
            return None
        if (len(items), widths.pop()) != (rows, cols):
            self._err("E-SHAPE", f"list is {len(items)}x{len(items[0])} but target is {target}",
                      st.line, value.col)
            return None
        return items

    def _clear(self, st: T.Clear) -> None:
        resolved = self._resolve_target(st.target, st.line)
        if resolved is None:
            return
        sheet, (c1, r1, c2, r2) = resolved
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.ops.append(ir.ClearCell(sheet, c, r, line=st.line))

    def _expect(self, st: T.Expect) -> None:
        host, _ = self._scope_sheet()
        if host is None and len(self.sheets) == 1:
            host = next(iter(self.sheets.values()))
        ast = self._bind_formula(st.formula, host, st.line)
        if ast is not None:
            self.ops.append(ir.Expect(host, ast, st.formula.text, st.line, st.formula.col))

    def _structural(self, st: T.Structural) -> None:
        sheet = self._sheet_of(None, st.line, st.col, f"{st.verb} {st.axis}")
        if sheet is None:
            return
        limit = MAX_ROW if st.axis == "rows" else MAX_COL
        if st.start + st.count - 1 > limit:
            self._err("E-REF", f"{st.axis} {st.text} are outside the sheet", st.line, st.col)
            return
        self.ops.append(ir.Structural(st.verb, st.axis, sheet, st.start, st.count, st.line))
        self.struct_line[sheet.casefold()] = st.line
        n = st.count if st.verb == "insert" else -st.count
        self.struct_ops.append((self.origin.get(sheet.casefold()) or sheet, st.axis, st.start, n))

    def _valid_new_sheet(self, name: str, line: int, col: int, renaming: str | None = None) -> bool:
        problem = None
        if not name.strip():
            problem = "sheet names cannot be empty"
        elif len(name) > 31:
            problem = "sheet names are at most 31 characters"
        elif set(name) & _BAD_SHEET_CHARS or name.startswith("'") or name.endswith("'"):
            problem = "sheet names cannot contain [ ] : * ? / \\ or start/end with '"
        elif name.casefold() in self.sheets and name.casefold() != (renaming or "").casefold():
            problem = f"a sheet named '{self.sheets[name.casefold()]}' already exists"
        if problem:
            self._err("E-SHEET", problem, line, col, len(name) + 2)
        return problem is None

    def _addsheet(self, st: T.AddSheet) -> None:
        if self._valid_new_sheet(st.name, st.line, st.col):
            self.sheets[st.name.casefold()] = st.name
            self.ops.append(ir.AddSheet(st.name, st.line))

    def _renamesheet(self, st: T.RenameSheet) -> None:
        old = self.sheets.get(st.old.casefold())
        if old is None:
            self._sheet_of(st.old, st.line, st.col_old + 1, "rename")
            return
        if not self._valid_new_sheet(st.new, st.line, st.col_new, renaming=old):
            return
        del self.sheets[old.casefold()]
        self.sheets[st.new.casefold()] = st.new
        self.origin[st.new.casefold()] = self.origin.pop(old.casefold(), None)
        for scope in self.scopes:
            if scope.sheet is not None and scope.sheet.casefold() == old.casefold():
                scope.sheet = st.new
            for sym in scope.symbols.values():
                if sym.kind == "range" and sym.sheet_cf == old.casefold():
                    sym.node = replace(sym.node, sheet=st.new)
                    sym.sheet_cf = st.new.casefold()
        self.ops.append(ir.RenameSheet(old, st.new, st.line))


def _text(node) -> str:
    from ..emitter import emit
    return emit(node, with_equals=False)
