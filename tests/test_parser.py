"""Parser tests: AST shape, precedence, associativity (task T1.7)."""
import pytest

from tabula.emitter import emit
from tabula.parser import (Binary, Bool, Call, Name, Number, ParseError, Percent,
                           RangeNode, Ref, Text, Unary, parse_formula)


def shape(node):
    """Fully parenthesized rendering, so precedence is directly assertable."""
    if isinstance(node, Number):
        whole = node.value == int(node.value)
        return str(int(node.value)) if whole else str(node.value)
    if isinstance(node, Text):
        return f'"{node.value}"'
    if isinstance(node, Bool):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, Ref):
        return (f"{node.sheet}!" if node.sheet else "") + str(node.ref)
    if isinstance(node, RangeNode):
        return (f"{node.sheet}!" if node.sheet else "") + str(node.rng)
    if isinstance(node, Name):
        return node.ident
    if isinstance(node, Unary):
        return f"(-{shape(node.operand)})"
    if isinstance(node, Percent):
        return f"({shape(node.operand)}%)"
    if isinstance(node, Binary):
        return f"({shape(node.left)} {node.op} {shape(node.right)})"
    if isinstance(node, Call):
        return f"{node.fname}({', '.join(shape(a) for a in node.args)})"
    return repr(node)


@pytest.mark.parametrize("src,expected", [
    ("2+3*4",        "(2 + (3 * 4))"),          # * binds tighter than +
    ("2*3+4",        "((2 * 3) + 4)"),
    ("1+2-3",        "((1 + 2) - 3)"),          # left-associative
    ("2^3^2",        "((2 ^ 3) ^ 2)"),          # ^ is LEFT-associative, as in Excel
    ("-A1^2",        "((-A1) ^ 2)"),            # unary binds tighter than ^
    ('"a"&1+2',      '("a" & (1 + 2))'),        # & is looser than +
    ("1+2<3*4",      "((1 + 2) < (3 * 4))"),    # comparison is loosest
    ("(2+3)*4",      "((2 + 3) * 4)"),
    ("50%*2",        "((50%) * 2)"),            # postfix % binds tighter than *
    ("2^50%",        "(2 ^ (50%))"),            # ... and tighter than ^
    ("-5%",          "((-5)%)"),                # negation binds tighter than %
])
def test_precedence_and_associativity(src, expected):
    assert shape(parse_formula(src)) == expected


def test_range_and_call():
    assert shape(parse_formula("SUM(A1:B4)")) == "SUM(A1:B4)"
    assert shape(parse_formula("IF(A1>0,1,2)")) == "IF((A1 > 0), 1, 2)"


def test_nested_calls():
    assert shape(parse_formula("MAX(SUM(A1:A3),AVERAGE(B1:B3))")) == \
        "MAX(SUM(A1:A3), AVERAGE(B1:B3))"


def test_excel_reference_forms():
    assert shape(parse_formula("SUM(Sales!E2:E21)")) == "SUM(Sales!E2:E21)"
    assert shape(parse_formula("'Q3 Sales'!A1*TaxRate")) == "(Q3 Sales!A1 * TaxRate)"
    assert shape(parse_formula("SUM(A:A)+SUM($2:$3)")) == "(SUM(A:A) + SUM($2:$3))"
    assert shape(parse_formula("_xlfn.XLOOKUP(1,A:A,B:B)")) == "XLOOKUP(1, A:A, B:B)"


@pytest.mark.parametrize("src", [
    "2^3^2", "(2^3)^2", "2^(3^2)", "1-(2-3)", "1-2-3", "-A1^2", "-(A1^2)", "50%*2",
    "(1+2)%", '"a""b"&TRUE', "SUM('Q3 Sales'!A1:B4,Sheet2!$C:$C,1:3)", "IF(A1<>\"\",#N/A,-1)",
    "TaxRate*Sales!#REF!", "1E+20/3",
])
def test_emitter_round_trip(src):
    """emit() is the inverse of parse(): parse(emit(parse(s))) == parse(s)."""
    ast = parse_formula(src)
    assert parse_formula(emit(ast)) == ast


@pytest.mark.parametrize("bad,message,pos", [
    ("SUM(A1:A3",  "expected ')' after argument list", 9),
    ("2+",         "unexpected end of formula",        2),
    ("(1+2",       "expected ')'",                     4),
    ("1 2",        "unexpected",                       2),
])
def test_syntax_errors_report_message_and_position(bad, message, pos):
    with pytest.raises(ParseError) as exc:
        parse_formula(bad)
    assert message in exc.value.message
    assert exc.value.pos == pos
