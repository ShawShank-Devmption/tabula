"""Diagnostics shared by the TEL compiler and the CLI (design.md section 2.2)."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class Diagnostic:
    severity: str          # 'error' | 'warning'
    code: str              # e.g. 'E-SHEET'
    message: str
    line: int = 0          # 1-based; 0 = not tied to a script line
    col: int = 0           # 1-based
    length: int = 1
    hint: str | None = None

    @property
    def is_error(self) -> bool:
        return self.severity == "error"

    def to_json(self) -> dict:
        out = {"severity": self.severity, "code": self.code, "line": self.line,
               "col": self.col, "message": self.message}
        if self.hint:
            out["hint"] = self.hint
        return out

    def render(self, source_lines: list[str] | None = None, filename: str = "script") -> str:
        where = f"{filename}:{self.line}:{self.col}: " if self.line else ""
        out = [f"{self.severity}[{self.code}] {where}{self.message}"]
        if self.line and source_lines and self.line <= len(source_lines):
            text = source_lines[self.line - 1].rstrip("\n")
            gutter = f"{self.line:>4} | "
            out.append(gutter + text)
            if self.col:
                out.append(" " * (len(gutter) - 2) + "| " + " " * (self.col - 1)
                           + "^" * max(1, self.length))
        if self.hint:
            out.append(f"  hint: {self.hint}")
        return "\n".join(out)


def error(code, message, line=0, col=0, length=1, hint=None) -> Diagnostic:
    return Diagnostic("error", code, message, line, col, length, hint)


def warning(code, message, line=0, col=0, length=1, hint=None) -> Diagnostic:
    return Diagnostic("warning", code, message, line, col, length, hint)


def levenshtein(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def suggest(name: str, candidates) -> str | None:
    """Closest candidate by edit distance (case-insensitive), if plausibly a typo."""
    best, best_d = None, None
    folded = name.casefold()
    for c in candidates:
        d = levenshtein(folded, c.casefold())
        if best_d is None or d < best_d or (d == best_d and c < best):
            best, best_d = c, d
    if best is None:
        return None
    limit = max(2, len(name) // 3)
    if best_d <= limit or folded in best.casefold() or best.casefold() in folded:
        return best
    return None
