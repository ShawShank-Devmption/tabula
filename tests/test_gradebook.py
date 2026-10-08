"""Integration test: the gradebook sample sheet (examples/gradebook.*).

Expected values are computed here independently of Tabula — plain Python
arithmetic over the same raw marks — so this is a real oracle, not a
restatement of whatever the engine happens to produce.  The same cell
definitions also generate examples/gradebook.xlsx, so a passing test means
the Excel workbook and the Tabula sheet agree.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from tabula.engine import Sheet
from tabula.values import display
from tools.make_gradebook import STUDENTS, build_cells


@pytest.fixture(scope="module")
def sheet() -> Sheet:
    s = Sheet()
    for ref, raw in build_cells():
        assert s.set(ref, raw) is None, f"{ref} failed to compile: {raw}"
    return s


# ---- independent expectations -------------------------------------------
TOTALS = {name: q1 + q2 + asg + exam for name, q1, q2, asg, exam in STUDENTS}
WEIGHTED = {name: round(q1 / 20 * 15 + q2 / 20 * 15 + asg / 20 * 20 + exam / 40 * 50, 1)
            for name, q1, q2, asg, exam in STUDENTS}


def letter(total: int) -> str:
    return "A" if total >= 90 else "B" if total >= 80 else "C" if total >= 70 else "F"


def test_totals_match_independent_arithmetic(sheet):
    for i, (name, *_) in enumerate(STUDENTS):
        assert display(sheet.value(f"F{i + 2}")) == str(TOTALS[name]), name


def test_grades_match_the_banding_rule(sheet):
    for i, (name, *_) in enumerate(STUDENTS):
        assert display(sheet.value(f"G{i + 2}")) == letter(TOTALS[name]), name


def test_weighted_percentages(sheet):
    for i, (name, *_) in enumerate(STUDENTS):
        expected = WEIGHTED[name]
        got = float(display(sheet.value(f"H{i + 2}")))
        assert got == pytest.approx(expected), name


def test_summary_block(sheet):
    totals = list(TOTALS.values())
    average = sum(totals) / len(totals)
    cases = {
        "B9":  str(round(average, 2)),
        "B10": str(max(totals)),
        "B11": str(min(totals)),
        "B12": str(len(totals)),
        "B13": str(max(totals) - min(totals)),
        "B14": "yes" if TOTALS["Aditi"] >= average else "no",
        "B15": f"Class average {round(average, 1)} of 100",
    }
    for ref, expected in cases.items():
        assert display(sheet.value(ref)) == expected, ref


def test_no_cell_holds_an_error(sheet):
    for ref, _ in build_cells():
        assert not display(sheet.value(ref)).startswith("#"), ref
