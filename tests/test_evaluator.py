"""Evaluation and error-value tests over the prototype grid (task T1.8)."""
import pytest

from tabula.engine import Sheet
from tabula.values import display


def sheet_with(**cells) -> Sheet:
    s = Sheet()
    for name, raw in cells.items():
        s.set(name, raw)
    return s


@pytest.mark.parametrize("formula,expected", [
    ("=2+3*4",        "14"),
    ("=2^3^2",        "64"),       # left-associative, as in Excel
    ("=(2+3)*4",      "20"),
    ("=-3^2",         "9"),        # (-3)^2 per unary-binds-tighter rule
    ("=7/2",          "3.5"),
    ("=MOD(7,3)",     "1"),
    ("=50%",          "0.5"),      # postfix percent
    ('="2"+1',        "3"),        # numeric text coerces, as in Excel
    ('="a"="A"',      "TRUE"),     # text comparison is case-insensitive
    ('=1<"a"',        "TRUE"),     # numbers sort before text
    ("=ROUND(2.5,0)", "3"),        # half away from zero (Python would give 2)
    ("=ROUND(2.675,2)", "2.68"),
    ("=1/3",          "0.333333333333333"),
    ('="a"&"b"',      "ab"),
    ('="n="&5',       "n=5"),      # number coerced to text by &
    ("=2<3",          "TRUE"),
    ("=2=2",          "TRUE"),
    ("=2<>2",         "FALSE"),
])
def test_expression_semantics(formula, expected):
    s = sheet_with(A1=formula)
    assert display(s.value("A1")) == expected


def test_references_and_aggregates():
    s = sheet_with(A1="10", A2="25", A3="40",
                   B1="=SUM(A1:A3)", B2="=AVERAGE(A1:A3)",
                   B3="=MIN(A1:A3)", B4="=MAX(A1:A3)", B5="=COUNT(A1:A3)")
    assert display(s.value("B1")) == "75"
    assert display(s.value("B2")) == "25"
    assert display(s.value("B3")) == "10"
    assert display(s.value("B4")) == "40"
    assert display(s.value("B5")) == "3"


def test_edit_propagates_to_dependents():
    s = sheet_with(A1="10", A2="25", B1="=SUM(A1:A2)*2")
    assert display(s.value("B1")) == "70"
    s.set("A2", "35")
    assert display(s.value("B1")) == "90"


def test_blank_cell_reads_as_zero_and_empty_text():
    s = sheet_with(A1="=B9+5", A2='=B9&"x"')
    assert display(s.value("A1")) == "5"
    assert display(s.value("A2")) == "x"


def test_if_evaluates_only_the_taken_branch():
    # If IF were eager, the untaken 1/0 branch would surface #DIV/0!.
    s = sheet_with(A1="=IF(TRUE, 42, 1/0)")
    assert display(s.value("A1")) == "42"


@pytest.mark.parametrize("formula,code", [
    ("=1/0",          "#DIV/0!"),
    ("=MOD(5,0)",     "#DIV/0!"),
    ('="abc"*2',      "#VALUE!"),   # non-numeric text in arithmetic
    ("=SQRT(-1)",     "#NUM!"),
    ("=NOSUCH(1)",    "#NAME?"),
])
def test_error_values(formula, code):
    s = sheet_with(A1=formula)
    assert display(s.value("A1")) == code


def test_errors_propagate_through_dependents():
    s = sheet_with(A1="=1/0", B1="=A1+1")
    assert display(s.value("B1")) == "#DIV/0!"


def test_circular_reference_is_detected():
    s = sheet_with(A1="=B1", B1="=A1")
    assert display(s.value("A1")) == "#CYCLE!"
    assert display(s.value("B1")) == "#CYCLE!"


def test_broken_formula_is_contained_to_its_cell():
    s = sheet_with(A1="10", B1="=SUM(A1", C1="=A1*2")
    assert display(s.value("B1")) == "#VALUE!"   # broken cell
    assert display(s.value("C1")) == "20"        # neighbours unaffected


def test_compile_error_reports_caret_diagnostic():
    s = Sheet()
    diag = s.set("C1", "=SUM(A1:A3")
    assert "error[SyntaxError] in C1" in diag
    assert "expected ')' after argument list" in diag
    assert diag.splitlines()[-1].strip() == "^"


def test_aggregates_skip_text_in_ranges_but_coerce_scalars():
    s = sheet_with(A1="10", A2="hello", A3="TRUE", B1="=SUM(A1:A3)", B2='=SUM(A1,"5",TRUE)',
                   B3="=COUNT(A1:A3)", B4="=COUNTA(A1:A3)")
    assert display(s.value("B1")) == "10"     # text and booleans in a range are skipped
    assert display(s.value("B2")) == "16"     # typed arguments are coerced
    assert display(s.value("B3")) == "1"
    assert display(s.value("B4")) == "3"


def test_countif_sumif_criteria():
    s = sheet_with(A1="North", A2="south", A3="North", B1="10", B2="20", B3="30",
                   C1='=SUMIF(A1:A3,"north",B1:B3)', C2='=COUNTIF(B1:B3,">15")',
                   C3='=COUNTIF(A1:A3,"N*")', C4='=SUMIF(B1:B3,"<>20")')
    assert display(s.value("C1")) == "40"     # case-insensitive match
    assert display(s.value("C2")) == "2"
    assert display(s.value("C3")) == "2"      # wildcard
    assert display(s.value("C4")) == "40"


def test_deep_chain_has_no_recursion_limit():
    s = Sheet()
    s.set("A1", "1")
    for r in range(2, 3001):
        s.set(f"A{r}", f"=A{r-1}+1")
    assert display(s.value("A3000")) == "3000"
