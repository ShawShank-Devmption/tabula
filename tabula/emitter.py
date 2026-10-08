"""Formula emitter: AST -> Excel formula text (the formula compiler's code generator).

Parentheses are inserted only where the binding powers require them, so
parse(emit(parse(s))) == parse(s) for every formula the parser accepts.
"""
from __future__ import annotations

from . import parser as P
from .functions import EXCEL_FUNCTIONS
from .lexer import TokType
from .refs import quote_sheet
from .values import format_number

_OP_BP = {
    "=": 10, "<>": 10, "<": 10, "<=": 10, ">": 10, ">=": 10,
    "&": 20, "+": 30, "-": 30, "*": 40, "/": 40, "^": 50,
}
_PERCENT_BP = 60
_UNARY_BP = 70
_ATOM_BP = 100

assert set(_OP_BP.values()) == set(P.BINARY_BP.values())  # one precedence table
assert all(t in P.BINARY_BP for t in (TokType.CARET, TokType.AMP))


def _bp(node) -> int:
    if isinstance(node, P.Binary):
        return _OP_BP[node.op]
    if isinstance(node, P.Percent):
        return _PERCENT_BP
    if isinstance(node, P.Unary):
        return _UNARY_BP
    return _ATOM_BP


def _wrap(node, needed: int, strict: bool, fx) -> str:
    text = _emit(node, fx)
    bp = _bp(node)
    return f"({text})" if (bp < needed or (strict and bp == needed)) else text


def _qualifier(sheet: str | None) -> str:
    return quote_sheet(sheet) + "!" if sheet else ""


def _call_name(fname: str, fx) -> str:
    """Name as stored in the file: Excel writes functions added after 2007 with
    '_xlfn.' (and FILTER/SORT also with '_xlws.'); without it Excel shows #NAME?."""
    if fx is None or fname in fx or fname not in EXCEL_FUNCTIONS:
        return fname
    return ("_xlfn._xlws." if fname in ("FILTER", "SORT") else "_xlfn.") + fname


def _emit(node, fx=None) -> str:
    if isinstance(node, P.Number):
        text = format_number(node.value)
        return text
    if isinstance(node, P.Text):
        return '"' + node.value.replace('"', '""') + '"'
    if isinstance(node, P.Bool):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, P.ErrorLit):
        return node.code
    if isinstance(node, P.Ref):
        return _qualifier(node.sheet) + str(node.ref)
    if isinstance(node, P.RangeNode):
        return _qualifier(node.sheet) + str(node.rng)
    if isinstance(node, P.Name):
        return _qualifier(node.sheet) + node.ident
    if isinstance(node, P.Unary):
        return node.op + _wrap(node.operand, _UNARY_BP, False, fx)
    if isinstance(node, P.Percent):
        return _wrap(node.operand, _PERCENT_BP, False, fx) + "%"
    if isinstance(node, P.Binary):
        bp = _OP_BP[node.op]
        return (_wrap(node.left, bp, False, fx) + node.op
                + _wrap(node.right, bp, True, fx))
    if isinstance(node, P.Call):
        return _call_name(node.fname, fx) + "(" + ",".join(_emit(a, fx) for a in node.args) + ")"
    raise TypeError(f"cannot emit {node!r}")


def emit(ast, with_equals: bool = True, plain_functions=None) -> str:
    """Excel formula text for an AST, e.g. '=SUM(Sales!E2:E21)*TaxRate'.

    plain_functions: when writing to a file, the set of functions Excel stores
    without a prefix (from openpyxl); every other known function gets '_xlfn.'.
    """
    return ("=" if with_equals else "") + _emit(ast, plain_functions)
