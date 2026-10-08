"""Recursive-descent parser with Pratt precedence climbing for formulas.

Grammar: design.md section 4.  Excel precedence, loosest to tightest:
comparison < '&' < '+ -' < '* /' < '^' (LEFT-associative) < postfix '%'
< prefix '-'/'+'.  So =2^3^2 is 64 and =-2^2 is 4, exactly as in Excel.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .lexer import Token, TokType, tokenize
from .refs import CellRef, RangeRef, parse_col, parse_row


# ---------- AST ----------
@dataclass
class Number:
    value: float


@dataclass
class Text:
    value: str


@dataclass
class Bool:
    value: bool


@dataclass
class ErrorLit:
    code: str


@dataclass
class Ref:
    ref: CellRef
    sheet: str | None = None  # qualifier as written; None = the host cell's sheet


@dataclass
class RangeNode:
    rng: RangeRef
    sheet: str | None = None


@dataclass
class Name:
    ident: str               # as written; resolution is case-insensitive
    sheet: str | None = None  # Sheet!Name (sheet-scoped defined name)


@dataclass
class Unary:
    op: str  # '-' or '+'
    operand: object


@dataclass
class Percent:
    operand: object


@dataclass
class Binary:
    op: str
    left: object
    right: object


@dataclass
class Call:
    fname: str
    args: list = field(default_factory=list)


class ParseError(Exception):
    def __init__(self, message: str, pos: int):
        super().__init__(message)
        self.message, self.pos = message, pos


# binding powers; every binary operator is left-associative in Excel
BINARY_BP = {
    TokType.EQ: 10, TokType.NE: 10, TokType.LT: 10, TokType.LE: 10,
    TokType.GT: 10, TokType.GE: 10,
    TokType.AMP: 20,
    TokType.PLUS: 30, TokType.MINUS: 30,
    TokType.STAR: 40, TokType.SLASH: 40,
    TokType.CARET: 50,
}


class Parser:
    def __init__(self, tokens: list[Token]):
        self.toks = tokens
        self.i = 0

    def peek(self, k: int = 0) -> Token:
        return self.toks[min(self.i + k, len(self.toks) - 1)]

    def advance(self) -> Token:
        tok = self.toks[self.i]
        self.i += 1
        return tok

    def expect(self, ttype: TokType, what: str) -> Token:
        if self.peek().type is ttype:
            return self.advance()
        raise ParseError(f"expected {what}", self.peek().pos)

    # -- Pratt loop -----------------------------------------------------
    def parse_expression(self, min_bp: int = 0):
        left = self.parse_postfix()
        while True:
            tok = self.peek()
            lbp = BINARY_BP.get(tok.type)
            if lbp is None or lbp < min_bp:
                return left
            self.advance()
            right = self.parse_expression(lbp + 1)  # +1: left-associative
            left = Binary(tok.lexeme, left, right)

    def parse_postfix(self):
        node = self.parse_prefix()
        while self.peek().type is TokType.PERCENT:
            self.advance()
            node = Percent(node)
        return node

    def parse_prefix(self):
        if self.peek().type in (TokType.MINUS, TokType.PLUS):
            tok = self.advance()
            return Unary(tok.lexeme, self.parse_prefix())
        return self.parse_primary()

    def parse_primary(self):
        tok = self.peek()
        if tok.type is TokType.NUMBER:
            if self.peek(1).type is TokType.COLON:      # row range 1:3
                return self.parse_reference(None)
            self.advance()
            return Number(tok.value)
        if tok.type is TokType.STRING:
            self.advance()
            return Text(tok.value)
        if tok.type is TokType.BOOL:
            self.advance()
            return Bool(tok.value)
        if tok.type is TokType.ERROR:
            self.advance()
            return ErrorLit(tok.value)
        if tok.type is TokType.SHEET:
            self.advance()
            return self.parse_reference(tok.value)
        if tok.type is TokType.IDENT and self.peek(1).type is TokType.LPAREN:
            self.advance()
            self.advance()
            args = []
            if self.peek().type is not TokType.RPAREN:
                args.append(self.parse_expression())
                while self.peek().type is TokType.COMMA:
                    self.advance()
                    args.append(self.parse_expression())
            self.expect(TokType.RPAREN, "')' after argument list")
            return Call(tok.value, args)
        if tok.type in (TokType.CELLREF, TokType.IDENT, TokType.COLPART, TokType.ROWPART):
            return self.parse_reference(None)
        if tok.type is TokType.LPAREN:
            self.advance()
            inner = self.parse_expression()
            self.expect(TokType.RPAREN, "')'")
            return inner
        what = f"unexpected {tok.lexeme!r}" if tok.lexeme else "unexpected end of formula"
        raise ParseError(what, tok.pos)

    # -- references: A1, A1:B4, A:C, 1:3, Name, (all optionally Sheet!-qualified)
    def parse_reference(self, sheet: str | None):
        tok = self.advance()
        nxt = self.peek()
        if tok.type is TokType.CELLREF:
            if nxt.type is TokType.COLON and self.peek(1).type is TokType.CELLREF:
                self.advance()
                end = self.advance()
                return RangeNode(RangeRef(tok.value, end.value), sheet)
            return Ref(tok.value, sheet)
        if tok.type in (TokType.IDENT, TokType.COLPART) and nxt.type is TokType.COLON:
            first = tok.value if tok.type is TokType.COLPART else parse_col(tok.lexeme)
            second_tok = self.peek(1)
            second = None
            if second_tok.type is TokType.COLPART:
                second = second_tok.value
            elif second_tok.type is TokType.IDENT:
                second = parse_col(second_tok.lexeme)
            if first is not None and second is not None:
                self.advance()
                self.advance()
                return RangeNode(RangeRef.columns(first[0], second[0], first[1], second[1]), sheet)
            raise ParseError("malformed column range", tok.pos)
        if tok.type in (TokType.NUMBER, TokType.ROWPART) and nxt.type is TokType.COLON:
            first = tok.value if tok.type is TokType.ROWPART else parse_row(tok.lexeme)
            second_tok = self.peek(1)
            second = None
            if second_tok.type is TokType.ROWPART:
                second = second_tok.value
            elif second_tok.type is TokType.NUMBER:
                second = parse_row(second_tok.lexeme)
            if first is not None and second is not None:
                self.advance()
                self.advance()
                return RangeNode(RangeRef.rows(first[0], second[0], first[1], second[1]), sheet)
            raise ParseError("malformed row range", tok.pos)
        if tok.type is TokType.IDENT:
            return Name(tok.lexeme, sheet)
        if tok.type is TokType.ERROR and sheet is not None:   # Sheet!#REF!
            return ErrorLit(tok.value)
        if tok.type in (TokType.COLPART, TokType.ROWPART):
            raise ParseError(f"'{tok.lexeme}' must be part of a range such as A:C or 1:3", tok.pos)
        raise ParseError("expected a reference after sheet name", tok.pos)


def parse_formula(text: str):
    """Parse formula text (leading '=' optional). Raises LexError/ParseError."""
    parser = Parser(tokenize(text))
    ast = parser.parse_expression()
    trailing = parser.peek()
    if trailing.type is not TokType.EOF:
        raise ParseError(f"unexpected {trailing.lexeme!r} after expression", trailing.pos)
    return ast


def parse_target(text: str):
    """Parse a bare reference or name (used for TEL targets and CLI arguments)."""
    parser = Parser(tokenize(text))
    tok = parser.peek()
    if tok.type is TokType.SHEET:
        parser.advance()
        node = parser.parse_reference(tok.value)
    elif tok.type in (TokType.CELLREF, TokType.IDENT, TokType.COLPART,
                      TokType.ROWPART, TokType.NUMBER):
        node = parser.parse_reference(None)
    else:
        raise ParseError("expected a cell, range or name", tok.pos)
    if not isinstance(node, (Ref, RangeNode, Name)):
        raise ParseError("expected a cell, range or name", tok.pos)
    trailing = parser.peek()
    if trailing.type is not TokType.EOF:
        raise ParseError(f"unexpected {trailing.lexeme!r} after reference", trailing.pos)
    return node


def walk(node):
    """Yield every node of an AST (pre-order)."""
    stack = [node]
    while stack:
        n = stack.pop()
        yield n
        if isinstance(n, (Unary, Percent)):
            stack.append(n.operand)
        elif isinstance(n, Binary):
            stack.extend((n.right, n.left))
        elif isinstance(n, Call):
            stack.extend(reversed(n.args))


# ---------- AST pretty-printer (explain) ----------
def _label(node) -> str:
    sheet = lambda n: f"{n.sheet}!" if getattr(n, "sheet", None) else ""
    if isinstance(node, Number):
        return f"Number({fmt_number(node.value)})"
    if isinstance(node, Text):
        return f'Text("{node.value}")'
    if isinstance(node, Bool):
        return f"Bool({node.value})"
    if isinstance(node, ErrorLit):
        return f"Error({node.code})"
    if isinstance(node, Ref):
        return f"Ref({sheet(node)}{node.ref})"
    if isinstance(node, RangeNode):
        return f"Range({sheet(node)}{node.rng})"
    if isinstance(node, Name):
        return f"Name({sheet(node)}{node.ident})"
    if isinstance(node, Unary):
        return f"Unary({node.op})"
    if isinstance(node, Percent):
        return "Percent(%)"
    if isinstance(node, Binary):
        return f"Binary({node.op})"
    return f"Call({node.fname})"


def dump_ast(node, prefix: str = "", is_last: bool = True,
             is_root: bool = True) -> list[str]:
    connector = "" if is_root else ("└── " if is_last else "├── ")
    lines = [prefix + connector + _label(node)]
    children = []
    if isinstance(node, (Unary, Percent)):
        children = [node.operand]
    elif isinstance(node, Binary):
        children = [node.left, node.right]
    elif isinstance(node, Call):
        children = node.args
    child_prefix = prefix + ("" if is_root else ("    " if is_last else "│   "))
    for k, child in enumerate(children):
        lines.extend(dump_ast(child, child_prefix, k == len(children) - 1, is_root=False))
    return lines


def fmt_number(v: float) -> str:
    return str(int(v)) if v == int(v) and abs(v) < 1e15 else repr(v)
