"""TEL scanner (design.md section 8.1).

TEL is line-oriented, and its lexical structure depends on position: after
`set TARGET`, an '=' switches the scanner into FORMULA mode and the rest of
the line becomes one token handed to the formula compiler (an "island"
grammar).  So instead of a free-running tokenizer, the parser drives a
LineScanner and asks for the token kind it expects next -- the lexer-mode
technique also used for regex-vs-division in JavaScript.
"""
from __future__ import annotations

import re

_WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_NUMBER = re.compile(r"-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")
_SPAN = re.compile(r"\$?[A-Za-z]{1,3}|\$?\d+")


class TelLexError(Exception):
    def __init__(self, message: str, col: int, length: int = 1, hint: str | None = None):
        super().__init__(message)
        self.message, self.col, self.length, self.hint = message, col, length, hint


class LineScanner:
    """Cursor over one source line; columns reported are 1-based."""

    def __init__(self, text: str):
        self.text = text.rstrip("\r\n")
        self.i = 0

    # -- basics ------------------------------------------------------------
    @property
    def col(self) -> int:
        return self.i + 1

    def skip_ws(self) -> None:
        while self.i < len(self.text) and self.text[self.i] in " \t":
            self.i += 1

    def at_end(self) -> bool:
        self.skip_ws()
        return self.i >= len(self.text)

    def peek(self) -> str:
        self.skip_ws()
        return self.text[self.i] if self.i < len(self.text) else ""

    def char(self, ch: str, what: str | None = None) -> int:
        self.skip_ws()
        col = self.col
        if self.peek() != ch:
            found = self.peek() or "end of line"
            raise TelLexError(f"expected {what or repr(ch)}, found {found!r}", col)
        self.i += 1
        return col

    # -- token kinds the parser can ask for ---------------------------------
    def word(self, what: str = "a word") -> tuple[str, int]:
        self.skip_ws()
        m = _WORD.match(self.text, self.i)
        if not m:
            found = self.peek() or "end of line"
            raise TelLexError(f"expected {what}, found {found!r}", self.col)
        col = self.col
        self.i = m.end()
        return m.group(0), col

    def string(self, what: str = "a quoted string") -> tuple[str, int]:
        self.skip_ws()
        col = self.col
        if self.peek() != '"':
            found = self.peek() or "end of line"
            raise TelLexError(f"expected {what}, found {found!r}", col)
        self.i += 1
        parts = []
        while True:
            if self.i >= len(self.text):
                raise TelLexError("unterminated string", col)
            ch = self.text[self.i]
            if ch == '"':
                if self.text[self.i + 1:self.i + 2] == '"':
                    parts.append('"')
                    self.i += 2
                    continue
                self.i += 1
                return "".join(parts), col
            parts.append(ch)
            self.i += 1

    def number(self) -> tuple[float, int] | None:
        self.skip_ws()
        m = _NUMBER.match(self.text, self.i)
        if not m:
            return None
        col = self.col
        value = float(m.group(0))
        if abs(value) > 1.7976931348623157e308:
            raise TelLexError("number is too large for Excel", col, m.end() - self.i)
        self.i = m.end()
        return value, col

    def target(self) -> tuple[str, int]:
        """A cell/range/name, possibly with a (quoted) sheet: one whitespace-free chunk."""
        self.skip_ws()
        col, start = self.col, self.i
        if self.peek() == "'":
            self.i += 1
            while self.i < len(self.text):
                if self.text[self.i] == "'":
                    if self.text[self.i + 1:self.i + 2] == "'":
                        self.i += 2
                        continue
                    break
                self.i += 1
            else:
                raise TelLexError("unterminated quoted sheet name", col)
            self.i += 1
        while self.i < len(self.text) and self.text[self.i] not in " \t=[\"":
            self.i += 1
        if self.i == start:
            found = self.peek() or "end of line"
            raise TelLexError(f"expected a cell, range or name, found {found!r}", col)
        return self.text[start:self.i], col

    def span(self) -> tuple[str, int]:
        """Row span '10' / '10:12' or column span 'C' / 'C:D'."""
        self.skip_ws()
        col = self.col
        m = _SPAN.match(self.text, self.i)
        if not m:
            found = self.peek() or "end of line"
            raise TelLexError(f"expected rows like 10 or 10:12, or columns like C or C:D, "
                              f"found {found!r}", col)
        end = m.end()
        if self.text[end:end + 1] == ":":
            m2 = _SPAN.match(self.text, end + 1)
            if not m2:
                raise TelLexError("incomplete span after ':'", col)
            end = m2.end()
        text = self.text[self.i:end]
        self.i = end
        return text, col

    def rest(self) -> tuple[str, int]:
        """FORMULA mode: everything up to the end of the line."""
        self.skip_ws()
        col = self.col
        text = self.text[self.i:].rstrip()
        self.i = len(self.text)
        return text, col

    def expect_end(self) -> None:
        if not self.at_end():
            raise TelLexError(f"unexpected {self.text[self.i:].strip()!r} at end of statement",
                              self.col, len(self.text[self.i:].strip()))
