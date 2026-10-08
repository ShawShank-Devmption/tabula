"""TEL parser: statements -> TEL AST (design.md section 8.2).

Recursive descent over lines, with newline-synchronised error recovery:
a malformed statement is reported and parsing resumes on the next line, so
one script run reports every independent mistake.  Formulas and targets are
parsed here by the formula compiler; their error positions are mapped back
into the script.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from ..diagnostics import Diagnostic, error
from ..lexer import LexError
from ..parser import ParseError, parse_formula, parse_target
from ..refs import parse_col, parse_row
from .lexer import LineScanner, TelLexError

KEYWORDS = {"in", "let", "set", "clear", "expect", "insert", "delete", "add", "rename"}


# ---------- TEL AST ----------
@dataclass
class TargetSrc:
    text: str
    col: int
    node: object  # formula AST: Ref | RangeNode | Name


@dataclass
class FormulaSrc:
    text: str
    col: int
    ast: object


@dataclass
class Literal:
    value: object
    col: int


@dataclass
class ListLit:
    rows: list  # list of rows (each a list of values); a 1-D list is one row
    one_dimensional: bool
    col: int


@dataclass
class InBlock:
    line: int
    sheet: str
    col: int
    body: list = field(default_factory=list)


@dataclass
class Let:
    line: int
    name: str
    col: int
    value: object  # TargetSrc | Literal


@dataclass
class Set:
    line: int
    target: TargetSrc
    value: object  # Literal | ListLit | FormulaSrc


@dataclass
class Clear:
    line: int
    target: TargetSrc


@dataclass
class Expect:
    line: int
    formula: FormulaSrc


@dataclass
class Structural:
    line: int
    verb: str   # 'insert' | 'delete'
    axis: str   # 'rows' | 'cols'
    start: int
    count: int
    col: int
    text: str


@dataclass
class AddSheet:
    line: int
    name: str
    col: int


@dataclass
class RenameSheet:
    line: int
    old: str
    new: str
    col_old: int
    col_new: int


@dataclass
class Script:
    statements: list
    diagnostics: list
    lines: list


class _Stmt(Exception):
    """Internal: a statement-level syntax error (carries a Diagnostic)."""

    def __init__(self, diag: Diagnostic):
        self.diag = diag


def parse_script(source: str) -> Script:
    lines = source.splitlines()
    diags: list[Diagnostic] = []
    root: list = []
    stack: list[InBlock] = []
    for number, text in enumerate(lines, 1):
        stripped = text.strip()
        if not stripped or stripped.startswith("#"):
            continue
        try:
            if stripped == "}":
                if not stack:
                    raise _Stmt(error("E-SYNTAX", "'}' without a matching 'in' block",
                                      number, text.index("}") + 1))
                block = stack.pop()
                (stack[-1].body if stack else root).append(block)
                continue
            stmt = _statement(LineScanner(text), number)
            if isinstance(stmt, InBlock):
                stack.append(stmt)
            else:
                (stack[-1].body if stack else root).append(stmt)
        except _Stmt as e:
            diags.append(e.diag)
        except TelLexError as e:
            code = "E-LEX" if "unterminated" in e.message else "E-SYNTAX"
            diags.append(error(code, e.message, number, e.col, e.length, e.hint))
    while stack:
        block = stack.pop()
        diags.append(error("E-SYNTAX", "'in' block opened here is never closed with '}'",
                           block.line, block.col))
        (stack[-1].body if stack else root).append(block)
    return Script(root, diags, lines)


def _statement(sc: LineScanner, line: int):
    keyword, kcol = sc.word("a statement keyword")
    kw = keyword.lower()
    if kw not in KEYWORDS:
        raise TelLexError(f"unknown statement '{keyword}'", kcol, len(keyword),
                          hint="statements: in, let, set, clear, expect, insert, delete, "
                               "add sheet, rename sheet")
    if kw == "in":
        if sc.peek() == '"':
            name, col = sc.string("a sheet name")
        else:
            name, col = sc.word("a sheet name")
        sc.char("{", "'{' to open the block")
        sc.expect_end()
        return InBlock(line, name, col)
    if kw == "let":
        name, col = sc.word("a name")
        sc.char("=", "'='")
        value = _literal(sc)
        if value is None:
            value = _target(sc, line)
        sc.expect_end()
        return Let(line, name, col, value)
    if kw == "set":
        target = _target(sc, line)
        if sc.peek() == "=":
            text, col = sc.rest()
            return Set(line, target, _formula(text, col, line))
        if sc.peek() == "[":
            value = _list(sc, line)
        else:
            value = _literal(sc)
            if value is None:
                found = sc.peek() or "end of line"
                raise TelLexError(f"expected a value (number, \"text\", TRUE/FALSE, [list] "
                                  f"or =formula), found {found!r}", sc.col)
        sc.expect_end()
        return Set(line, target, value)
    if kw == "clear":
        target = _target(sc, line)
        sc.expect_end()
        return Clear(line, target)
    if kw == "expect":
        text, col = sc.rest()
        if not text:
            raise TelLexError("expected a condition after 'expect'", col)
        return Expect(line, _formula(text, col, line))
    if kw in ("insert", "delete"):
        axis, acol = sc.word("'rows' or 'cols'")
        axis = axis.lower()
        if axis in ("row", "column", "columns", "col"):
            axis = "rows" if axis == "row" else "cols"
        if axis not in ("rows", "cols"):
            raise TelLexError(f"expected 'rows' or 'cols', found '{axis}'", acol)
        text, col = sc.span()
        sc.expect_end()
        start, count = _span_bounds(text, axis, col, line)
        return Structural(line, kw, axis, start, count, col, text)
    if kw == "add":
        what, wcol = sc.word("'sheet'")
        if what.lower() != "sheet":
            raise TelLexError("expected 'sheet' after 'add'", wcol)
        name, col = sc.string("a sheet name in quotes")
        sc.expect_end()
        return AddSheet(line, name, col)
    # rename sheet "Old" to "New"
    what, wcol = sc.word("'sheet'")
    if what.lower() != "sheet":
        raise TelLexError("expected 'sheet' after 'rename'", wcol)
    old, c1 = sc.string("the current sheet name in quotes")
    to, tcol = sc.word("'to'")
    if to.lower() != "to":
        raise TelLexError("expected 'to'", tcol)
    new, c2 = sc.string("the new sheet name in quotes")
    sc.expect_end()
    return RenameSheet(line, old, new, c1, c2)


def _literal(sc: LineScanner):
    ch = sc.peek()
    if ch == '"':
        value, col = sc.string()
        return Literal(value, col)
    num = sc.number()
    if num is not None:
        return Literal(num[0], num[1])
    save = sc.i
    try:
        word, col = sc.word()
    except TelLexError:
        return None
    if word.upper() in ("TRUE", "FALSE"):
        return Literal(word.upper() == "TRUE", col)
    sc.i = save
    return None


def _list(sc: LineScanner, line: int) -> ListLit:
    col = sc.char("[")
    items = []
    nested = sc.peek() == "["
    while True:
        if nested:
            inner = _list(sc, line)
            if not inner.one_dimensional:
                raise TelLexError("lists nest at most two levels", inner.col)
            items.append(inner.rows[0])
        else:
            lit = _literal(sc)
            if lit is None:
                found = sc.peek() or "end of line"
                raise TelLexError(f"expected a list item, found {found!r}", sc.col)
            items.append(lit.value)
        if sc.peek() == ",":
            sc.i += 1
            continue
        sc.char("]", "',' or ']'")
        break
    if nested:
        return ListLit(items, False, col)
    return ListLit([items], True, col)


def _target(sc: LineScanner, line: int) -> TargetSrc:
    text, col = sc.target()
    try:
        node = parse_target(text)
    except (LexError, ParseError) as e:
        raise _Stmt(error("E-SYNTAX", f"bad target '{text}': {e.message}", line, col + e.pos,
                          hint="targets look like A1, A1:C9, Sales!B2 or 'Q3 Sales'!A1:B4"))
    return TargetSrc(text, col, node)


def _formula(text: str, col: int, line: int) -> FormulaSrc:
    try:
        ast = parse_formula(text)
    except LexError as e:
        raise _Stmt(error("E-LEX", f"in formula: {e.message}", line, col + e.pos))
    except ParseError as e:
        raise _Stmt(error("E-SYNTAX", f"in formula: {e.message}", line, col + e.pos))
    return FormulaSrc(text, col, ast)


def _span_bounds(text: str, axis: str, col: int, line: int) -> tuple[int, int]:
    parts = text.replace("$", "").split(":")
    parse = parse_row if axis == "rows" else parse_col
    try:
        nums = [parse(p)[0] for p in parts]
    except TypeError:
        kind = "row numbers like 10:12" if axis == "rows" else "column letters like C:D"
        raise _Stmt(error("E-SYNTAX", f"'{text}' is not a valid {axis} span; use {kind}",
                          line, col, len(text)))
    start, end = min(nums), max(nums)
    return start, end - start + 1
