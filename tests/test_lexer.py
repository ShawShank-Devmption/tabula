"""Unit tests for the formula scanner (task T1.6)."""
import pytest

from tabula.lexer import LexError, TokType, tokenize


def types(text):
    return [t.type for t in tokenize(text) if t.type is not TokType.EOF]


def test_every_token_class():
    toks = tokenize('SUM(A1:$B$2) + 3.5 & "hi" <> TRUE')
    assert [t.type for t in toks] == [
        TokType.IDENT, TokType.LPAREN, TokType.CELLREF, TokType.COLON,
        TokType.CELLREF, TokType.RPAREN, TokType.PLUS, TokType.NUMBER,
        TokType.AMP, TokType.STRING, TokType.NE, TokType.BOOL, TokType.EOF,
    ]


@pytest.mark.parametrize("text,value", [("12", 12.0), ("3.5", 3.5), (".5", 0.5)])
def test_number_forms(text, value):
    assert tokenize(text)[0].value == value


def test_string_escape_is_doubled_quote():
    assert tokenize('"say ""hi"""')[0].value == 'say "hi"'


def test_absolute_reference_keeps_anchors():
    ref = tokenize("$B$2")[0].value
    assert (ref.col, ref.row, ref.abs_col, ref.abs_row) == (2, 2, True, True)


def test_two_char_operators_win_over_one_char():
    assert types("a<=b") == [TokType.IDENT, TokType.LE, TokType.IDENT]
    assert types("a<>b") == [TokType.IDENT, TokType.NE, TokType.IDENT]
    assert types("a<b") == [TokType.IDENT, TokType.LT, TokType.IDENT]


def test_identifier_vs_reference_classification():
    assert types("A1") == [TokType.CELLREF]
    assert types("A1X") == [TokType.IDENT]     # not a reference shape
    assert types("XFD1048576") == [TokType.CELLREF]   # Excel's last cell
    assert types("XFE1") == [TokType.IDENT]           # past column XFD -> not a ref
    assert types("A1048577") == [TokType.IDENT]       # past the last row -> not a ref


def test_sheet_qualifiers():
    toks = tokenize("Sales!E2+'Q3 Sales'!A1+'It''s'!B2+Q3!C3")
    sheets = [t.value for t in toks if t.type is TokType.SHEET]
    assert sheets == ["Sales", "Q3 Sales", "It's", "Q3"]   # Q3! is a sheet, not a cell


def test_error_literals_and_storage_prefix():
    assert [t.value for t in tokenize("#REF!+#DIV/0!") if t.type is TokType.ERROR] == \
        ["#REF!", "#DIV/0!"]
    assert tokenize("_xlfn.XLOOKUP(1)")[0].value == "XLOOKUP"


def test_column_and_row_parts():
    assert types("$A:$C") == [TokType.COLPART, TokType.COLON, TokType.COLPART]
    assert types("$1:$3") == [TokType.ROWPART, TokType.COLON, TokType.ROWPART]
    assert tokenize("1.5E3")[0].value == 1500.0


def test_positions_are_recorded():
    toks = tokenize("1 + 23")
    assert [t.pos for t in toks[:3]] == [0, 2, 4]


@pytest.mark.parametrize("bad,message", [
    ('"unterminated', "unterminated string literal"),
    ("1 # 2", "unknown error literal"),
    ("$1A", "malformed cell reference"),
    ("{1,2}", "array constants"),
    ("'Sheet", "unterminated quoted sheet name"),
])
def test_lexical_errors(bad, message):
    with pytest.raises(LexError) as exc:
        tokenize(bad)
    assert message in exc.value.message
