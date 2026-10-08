"""Hand-written scanner for the Excel-compatible formula language.

Tokenizes formula text (with or without the leading '=').  Every token
carries its start column so diagnostics can point at the exact spot.
Classification rules are in design.md section 3.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto

from .refs import CellRef, parse_col, parse_row
from .values import ERROR_CODES


class TokType(Enum):
    NUMBER = auto()
    STRING = auto()
    BOOL = auto()
    ERROR = auto()      # #REF!, #DIV/0!, ...
    CELLREF = auto()    # A1, $B$2
    COLPART = auto()    # $A (only meaningful inside A:C)
    ROWPART = auto()    # $1 (only meaningful inside 1:3)
    SHEET = auto()      # Sales!  or  'Q3 Sales'!
    IDENT = auto()      # function or defined name (also bare column letters)
    PLUS = auto(); MINUS = auto(); STAR = auto(); SLASH = auto()
    PERCENT = auto(); CARET = auto(); AMP = auto()
    EQ = auto(); NE = auto(); LT = auto(); LE = auto(); GT = auto(); GE = auto()
    LPAREN = auto(); RPAREN = auto(); COMMA = auto(); COLON = auto()
    EOF = auto()


@dataclass
class Token:
    type: TokType
    lexeme: str
    value: object  # parsed literal / CellRef / name where applicable
    pos: int       # 0-based start column in the formula text


class LexError(Exception):
    def __init__(self, message: str, pos: int):
        super().__init__(message)
        self.message, self.pos = message, pos


_SINGLE = {
    "+": TokType.PLUS, "-": TokType.MINUS, "*": TokType.STAR,
    "/": TokType.SLASH, "%": TokType.PERCENT, "^": TokType.CARET,
    "&": TokType.AMP, "=": TokType.EQ, "(": TokType.LPAREN,
    ")": TokType.RPAREN, ",": TokType.COMMA, ":": TokType.COLON,
}
# Prefixes Excel uses when *storing* newer functions in the file.
_STORAGE_PREFIXES = ("_XLFN.", "_XLWS.", "_XLPM.")
_WORD_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.$\\")
_ERRORS_LONGEST_FIRST = sorted(ERROR_CODES, key=len, reverse=True)


def _scan_number(text: str, i: int) -> int:
    n = len(text)
    while i < n and text[i].isdigit():
        i += 1
    if i < n and text[i] == ".":
        i += 1
        while i < n and text[i].isdigit():
            i += 1
    if i < n and text[i] in "eE":
        j = i + 1
        if j < n and text[j] in "+-":
            j += 1
        if j < n and text[j].isdigit():
            i = j
            while i < n and text[i].isdigit():
                i += 1
    return i


def tokenize(text: str) -> list[Token]:
    toks: list[Token] = []
    i, n = 0, len(text)
    if text.startswith("="):
        i = 1
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
            continue
        start = i
        # -- quoted sheet qualifier: 'Q3 Sales'!
        if ch == "'":
            i, parts = i + 1, []
            while True:
                if i >= n:
                    raise LexError("unterminated quoted sheet name", start)
                if text[i] == "'":
                    if i + 1 < n and text[i + 1] == "'":
                        parts.append("'")
                        i += 2
                        continue
                    i += 1
                    break
                parts.append(text[i])
                i += 1
            if i >= n or text[i] != "!":
                raise LexError("quoted sheet name must be followed by '!'", start)
            i += 1
            toks.append(Token(TokType.SHEET, text[start:i], "".join(parts), start))
            continue
        # -- error literals
        if ch == "#":
            for code in _ERRORS_LONGEST_FIRST:
                if text[i:i + len(code)].upper() == code:
                    toks.append(Token(TokType.ERROR, text[i:i + len(code)], code, start))
                    i += len(code)
                    break
            else:
                raise LexError("unknown error literal or unsupported '#' operator", start)
            continue
        # -- numbers: 12, 3.5, .5, 1E3
        if ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit()):
            i = _scan_number(text, i)
            lexeme = text[start:i]
            toks.append(Token(TokType.NUMBER, lexeme, float(lexeme), start))
            continue
        # -- strings: "..." with "" as the escaped quote
        if ch == '"':
            i, parts = i + 1, []
            while True:
                if i >= n:
                    raise LexError("unterminated string literal", start)
                if text[i] == '"':
                    if i + 1 < n and text[i + 1] == '"':
                        parts.append('"')
                        i += 2
                        continue
                    i += 1
                    break
                parts.append(text[i])
                i += 1
            toks.append(Token(TokType.STRING, text[start:i], "".join(parts), start))
            continue
        # -- words: references, sheet qualifiers, booleans, names, $A / $1 parts
        if ch.isalpha() or ch in "_\\$":
            while i < n and text[i] in _WORD_CHARS:
                i += 1
            word = text[start:i]
            if i < n and text[i] == "!" and "$" not in word:
                i += 1
                toks.append(Token(TokType.SHEET, text[start:i], word, start))
                continue
            ref = CellRef.parse(word)
            upper = word.upper()
            if ref is not None:
                toks.append(Token(TokType.CELLREF, word, ref, start))
            elif upper in ("TRUE", "FALSE"):
                toks.append(Token(TokType.BOOL, word, upper == "TRUE", start))
            elif "$" in word:
                col = parse_col(word)
                row = parse_row(word)
                if col is not None:
                    toks.append(Token(TokType.COLPART, word, col, start))
                elif row is not None:
                    toks.append(Token(TokType.ROWPART, word, row, start))
                else:
                    raise LexError(f"malformed cell reference '{word}'", start)
            else:
                stripped = True
                while stripped:
                    stripped = False
                    for prefix in _STORAGE_PREFIXES:
                        if upper.startswith(prefix):
                            upper, stripped = upper[len(prefix):], True
                toks.append(Token(TokType.IDENT, word, upper, start))
            continue
        # -- two-char operators before one-char (maximal munch)
        two = text[i:i + 2]
        if two in ("<>", "<=", ">="):
            ttype = {"<>": TokType.NE, "<=": TokType.LE, ">=": TokType.GE}[two]
            toks.append(Token(ttype, two, None, i))
            i += 2
            continue
        if ch == "<":
            toks.append(Token(TokType.LT, ch, None, i)); i += 1; continue
        if ch == ">":
            toks.append(Token(TokType.GT, ch, None, i)); i += 1; continue
        if ch in _SINGLE:
            toks.append(Token(_SINGLE[ch], ch, None, i)); i += 1; continue
        if ch == "{":
            raise LexError("array constants {...} are not supported", i)
        raise LexError(f"illegal character {ch!r}", i)
    toks.append(Token(TokType.EOF, "", None, n))
    return toks
