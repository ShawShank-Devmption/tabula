"""Value model and Excel-compatible coercion rules (design.md section 4.2).

Values are plain Python objects: float (numbers), str (text), bool,
None (blank) and ErrorValue.  Every rule that turns one type into another
lives here so the evaluator never improvises.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class ErrorValue:
    code: str  # '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#NULL!', '#CYCLE!'
    detail: str = ""

    def __eq__(self, other) -> bool:  # errors compare by code only
        return isinstance(other, ErrorValue) and other.code == self.code

    def __hash__(self) -> int:
        return hash(self.code)

    def __str__(self) -> str:
        return self.code


DIV0 = ErrorValue("#DIV/0!")
VALUE_ERR = ErrorValue("#VALUE!")
REF_ERR = ErrorValue("#REF!")
NAME_ERR = ErrorValue("#NAME?")
NA_ERR = ErrorValue("#N/A")
NUM_ERR = ErrorValue("#NUM!")
NULL_ERR = ErrorValue("#NULL!")
CYCLE_ERR = ErrorValue("#CYCLE!")  # simulator-internal; never written to a file

ERROR_CODES = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!")

BLANK = None  # an empty cell reads as Blank


class _Unknown:
    """Value of a cell Tabula cannot compute (unsimulated, inputs changed)."""

    def __repr__(self) -> str:
        return "UNKNOWN"


UNKNOWN = _Unknown()

_NUMERIC_TEXT = re.compile(r"^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*(%?)\s*$")


def is_error(v) -> bool:
    return isinstance(v, ErrorValue)


def is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def error_from_code(code: str) -> ErrorValue:
    return ErrorValue(code.upper())


def to_number(v):
    """Coerce for arithmetic: blank->0, bool->1/0, numeric text->number."""
    if is_error(v):
        return v
    if v is BLANK:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if is_number(v):
        return float(v)
    if isinstance(v, str):
        m = _NUMERIC_TEXT.match(v)
        if m:
            n = float(v.strip().rstrip("%").strip())
            if not math.isfinite(n):
                return NUM_ERR
            return n / 100 if m.group(3) else n
    return VALUE_ERR


def format_number(v: float) -> str:
    """Excel 'General' rendering: integers without '.0', 15 significant digits."""
    if not math.isfinite(v):
        return "#NUM!"
    if v == int(v) and abs(v) < 1e15:
        return str(int(v))
    return f"{v:.15g}"


def to_text(v):
    if is_error(v):
        return v
    if v is BLANK:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if is_number(v):
        return format_number(float(v))
    return str(v)


def to_bool(v):
    """Coerce for logical context (IF condition, AND/OR arguments)."""
    if is_error(v):
        return v
    if v is BLANK:
        return False
    if isinstance(v, bool):
        return v
    if is_number(v):
        return v != 0
    if isinstance(v, str) and v.strip().upper() in ("TRUE", "FALSE"):
        return v.strip().upper() == "TRUE"
    return VALUE_ERR


def _type_rank(v) -> int:
    """Excel orders values of different types: numbers < text < booleans."""
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def compare(a, b) -> int:
    """-1/0/1 per Excel comparison rules. Callers handle errors first."""
    # A blank takes the type of the other operand.
    if a is BLANK and b is BLANK:
        return 0
    if a is BLANK:
        a = False if isinstance(b, bool) else "" if isinstance(b, str) else 0.0
    if b is BLANK:
        b = False if isinstance(a, bool) else "" if isinstance(a, str) else 0.0
    ra, rb = _type_rank(a), _type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if isinstance(a, str):
        a, b = a.casefold(), b.casefold()
    elif ra == 0:
        a, b = _sig15(a), _sig15(b)  # Excel compares numbers to 15 significant digits
    return (a > b) - (a < b)


def _sig15(x) -> float:
    return float(f"{x:.15g}")


def display(v) -> str:
    """Human-readable form for grids and plans."""
    if v is BLANK:
        return ""
    if v is UNKNOWN:
        return "?"
    if is_error(v):
        return v.code
    return to_text(v)


def values_equal(a, b) -> bool:
    """Equality for comparing engine results with Excel's cached values."""
    if is_number(a) and is_number(b):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    if a in (BLANK, "") and b in (BLANK, ""):
        return True
    if type(a) is not type(b):
        return False
    return a == b
