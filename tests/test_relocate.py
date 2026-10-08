"""Relocation tests: fill shifting, insert/delete rows/cols, sheet rename (design.md 5.5).

Every expected result is Excel's behaviour for the same edit.
"""
import pytest

from tabula.emitter import emit
from tabula.parser import parse_formula
from tabula.relocate import FillError, relocate, rename_sheet, shift_for_fill


def reloc(formula, at, n, axis="rows", host="Sales", sheet="Sales"):
    ast, _ = relocate(parse_formula(formula), host, sheet, axis, at, n)
    return emit(ast)


@pytest.mark.parametrize("formula,expected", [
    ("=SUM(E2:E21)",   "=SUM(E2:E22)"),     # insertion inside the range: expands
    ("=E22",           "=E23"),             # below: moves
    ("=$E$22",         "=$E$23"),           # absolute references move too
    ("=E5",            "=E5"),              # above: unchanged
    ("=SUM(E12:E21)",  "=SUM(E13:E22)"),    # range starting at the insertion: moves
    ("=SUM(E2:E10)",   "=SUM(E2:E10)"),     # range ending just above: does NOT expand (Excel)
    ("=SUM(E:E)",      "=SUM(E:E)"),        # whole columns are unaffected by row edits
    ("=SUM(5:30)",     "=SUM(5:31)"),       # whole rows expand like ranges
])
def test_insert_one_row_at_11(formula, expected):
    assert reloc(formula, 11, 1) == expected


@pytest.mark.parametrize("formula,expected", [
    ("=SUM(E2:E21)",   "=SUM(E2:E20)"),     # one row deleted inside: shrinks
    ("=E8",            "=#REF!"),           # the deleted cell itself
    ("=E22",           "=E21"),             # below: moves up
    ("=SUM(E8:E8)",    "=SUM(#REF!)"),      # a range wholly deleted
    ("=SUM(E5:E8)",    "=SUM(E5:E7)"),      # deletion at the range's end
    ("=SUM(E8:E12)",   "=SUM(E8:E11)"),     # deletion at the range's start
])
def test_delete_row_8(formula, expected):
    assert reloc(formula, 8, -1) == expected


def test_delete_span_and_columns():
    assert reloc("=SUM(E2:E21)", 20, -3) == "=SUM(E2:E19)"
    assert reloc("=SUM(C2:G2)", 4, 2, axis="cols") == "=SUM(C2:I2)"     # insert at column D
    assert reloc("=H2", 4, -2, axis="cols") == "=F2"                     # delete D:E


def test_only_references_to_the_edited_sheet_move():
    assert reloc("=Sales!E22+Inputs!B22+E22", 11, 1, host="Summary") == \
        "=Sales!E23+Inputs!B22+E22"   # unqualified E22 on Summary is Summary!E22


def test_fill_respects_dollar_anchors():
    ast = parse_formula("=E2*$B$1+SUM($C2:C$2)")
    shifted, changed = shift_for_fill(ast, 1, 3)
    assert changed and emit(shifted) == "=F5*$B$1+SUM($C5:D$2)"


def test_fill_off_the_grid_is_an_error():
    with pytest.raises(FillError):
        shift_for_fill(parse_formula("=A2"), -1, 0)


def test_rename_quotes_names_that_need_it():
    ast, changed = rename_sheet(parse_formula("=Sales!E22+sales!F2+Summary!B1"), "Sales", "Q3 Sales")
    assert changed and emit(ast) == "='Q3 Sales'!E22+'Q3 Sales'!F2+Summary!B1"
