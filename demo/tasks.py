"""The five demo tasks: what the agent is asked, and how the result is judged.

Prompts are given verbatim to the agent in both arms.  Expected values are
computed here in plain Python from build_workbook.DATA -- independently of
Tabula and of any formula engine.  Checks are run by evaluate.py, which
evaluates the agent's output workbook with pycel (an independent formula
engine).  Standard library only, so both the runner and the checker import it.
"""
from __future__ import annotations

from dataclasses import dataclass

from build_workbook import COMMISSION_RATE, DATA, TARGET, TAX_RATE


@dataclass(frozen=True)
class Task:
    id: str
    title: str
    workbook: str          # file in demo/workbooks/
    prompt: str
    what_breaks: str       # the trap, for the report


TASKS = [
    Task("t1_tax_column", "Add a computed column", "sales_q3.xlsx",
         'In the Sales sheet, add a new column G with the header "Tax" in G1. For every sales '
         "row, Tax is the row's Revenue multiplied by the tax rate on the Inputs sheet. Put the "
         "column total in the Total row (row 22). Then, on the Summary sheet, add a line in row 8: "
         'the label "Total tax" in A8 and the total tax in B8.',
         "baseline: no structural change; both methods should manage"),
    Task("t2_insert_row", "Insert a missed sale", "sales_q3.xlsx",
         "A sale was left out of the Sales sheet. Insert it as a new row directly below row 10 "
         '(so it becomes row 11): Region "West", Rep "Dana Lee", Units 40, Unit Price 125, with '
         "Revenue and Commission calculated the same way as the other rows. Every total and "
         "every figure on the Summary sheet must include the new sale.",
         "openpyxl insert_rows moves cells but rewrites no formulas or defined names"),
    Task("t3_delete_test_rows", "Delete test rows", "sales_q3.xlsx",
         'The Sales sheet contains two test entries (Rep "TEST"). Delete those two rows entirely '
         "so the rows below move up. All totals and all Summary figures must remain correct.",
         "openpyxl delete_rows leaves totals summing stale ranges (even themselves: a cycle)"),
    Task("t4_fix_summary", "Fix two wrong figures", "sales_q3_buggy.xlsx",
         "Two of the figures on the Summary sheet are calculated wrongly. Find and fix them so "
         "that every Summary figure is correct. Do not change anything on the other sheets.",
         "debugging: needs inspection of formulas, not just values"),
    Task("t5_rename_sheet", "Rename a sheet", "sales_q3.xlsx",
         'Rename the sheet "Sales" to "Q3 Sales". Everything that refers to it (formulas on other '
         "sheets and named ranges) must keep working.",
         "openpyxl renames the tab only; references to 'Sales' break, and the new name needs quotes"),
]
TASKS_BY_ID = {t.id: t for t in TASKS}


# ---------------------------------------------------------------- expectations
def rows_after(task_id: str) -> list[tuple]:
    """Data rows (region, rep, units, price) the Sales sheet should hold after the task."""
    data = list(DATA)
    if task_id == "t2_insert_row":
        data.insert(9, ("West", "Dana Lee", 40, 125))       # sheet row 11 = data index 9
    elif task_id == "t3_delete_test_rows":
        data = [d for d in data if d[1] != "TEST"]
    return data


def figures(data: list[tuple]) -> dict:
    revenue = [u * p for _, _, u, p in data]
    total = sum(revenue)
    return {
        "units": sum(u for _, _, u, _ in data),
        "revenue": revenue,
        "total_revenue": total,
        "total_commission": total * COMMISSION_RATE,
        "average": total / len(data),
        "north": sum(u * p for reg, _, u, p in data if reg == "North"),
        "target_met": "Yes" if total >= TARGET else "No",
        "total_tax": total * TAX_RATE,
    }
