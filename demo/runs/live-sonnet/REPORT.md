# Tabula vs default — A/B agent demo

- run folder: `live-sonnet`  ·  generated 2026-10-08 03:32
- driver: **claude**  ·  model: `claude-sonnet-5-5`
- checker: independent formula engine (pycel) + plain-Python expected values; Tabula's engine is not used to judge Tabula

## Summary

| arm | tasks fully correct | checks passed | silent failures | avg cost (USD) | avg time (s) | avg tool calls |
|---|---|---|---|---|---|---|
| default (Python + openpyxl) | 11/15 | 129/141 | 4 | 0.029 | 11 | 2.5 |
| Tabula (TEL) | 15/15 | 141/141 | 0 | 0.036 | 11 | 4.8 |

*Silent failure*: the agent finished normally (no error, no timeout) but the workbook fails the checks — the failure an unchecked edit hides.

## Per task

| task | trap | default (Python + openpyxl) | Tabula (TEL) |
|---|---|---|---|
| Add a computed column | baseline: no structural change; both methods should manage | ✅ 11/11 ✅ 11/11 ✅ 11/11 | ✅ 11/11 ✅ 11/11 ✅ 11/11 |
| Insert a missed sale | openpyxl insert_rows moves cells but rewrites no formulas or defined names | ❌ 7/11 ✅ 11/11 ❌ 7/11 | ✅ 11/11 ✅ 11/11 ✅ 11/11 |
| Delete test rows | openpyxl delete_rows leaves totals summing stale ranges (even themselves: a cycle) | ❌ 7/9 ✅ 9/9 ❌ 7/9 | ✅ 9/9 ✅ 9/9 ✅ 9/9 |
| Fix two wrong figures | debugging: needs inspection of formulas, not just values | ✅ 9/9 ✅ 9/9 ✅ 9/9 | ✅ 9/9 ✅ 9/9 ✅ 9/9 |
| Rename a sheet | openpyxl renames the tab only; references to 'Sales' break, and the new name needs quotes | ✅ 7/7 ✅ 7/7 ✅ 7/7 | ✅ 7/7 ✅ 7/7 ✅ 7/7 |

## Runs

### Add a computed column — default (Python + openpyxl) (trial 1): ✅ correct  ·  $0.026 · 9s · 2 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column to the Sales sheet and a Total tax line to the Summary sheet, and saved both in `book.xlsx`. I haven't opened the file in Excel or recalculated it, so the formulas haven't been checked for computed values. - **Sales sheet:** - G1 is the header "Tax". - G2:G21 each hold `=E{row}*TaxRate`. `TaxRate` is the existing named range pointing to Inputs!B2 (18%). - G22 is `=SUM(G2:G21)`. - I copied the formatting from column F, including the column width. - **Summary sheet:** A8 is "Total tax" and B8 is `=Sales!G22`. Both are formatted like row 4.

files: `demo/runs/live-sonnet/t1_tax_column__default__1/` (book.xlsx, transcript)

### Add a computed column — default (Python + openpyxl) (trial 2): ✅ correct  ·  $0.031 · 9s · 2 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column and the Summary line, and saved both to `book.xlsx`. I haven't opened the file in Excel or recalculated it, so the new formulas haven't been checked. - **Sales sheet:** G1 now says "Tax". G2:G21 each hold `=E{row}*TaxRate`, which uses the named range pointing to the tax rate on the Inputs sheet (18%). G22 holds `=SUM(G2:G21)`. I copied the formatting from column F. - **Summary sheet:** A8 says "Total tax" and B8 holds `=Sales!G22`. Rows 8 and 15 on the Sales sheet are test rows (Region "Test", Rep "TEST"). Each has 1 unit at a price of 1. I gave them tax like every other…

files: `demo/runs/live-sonnet/t1_tax_column__default__2/` (book.xlsx, transcript)

### Add a computed column — default (Python + openpyxl) (trial 3): ✅ correct  ·  $0.031 · 10s · 3 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column to the Sales sheet and a Total tax line to the Summary sheet, and saved `book.xlsx` in place. I didn't recalculate the workbook or check the computed values. - **Sales:** G1 is the header "Tax". G2:G21 are `=E{row}*TaxRate`, using the existing named range for the Inputs tax rate (18%). G22 is `=SUM(G2:G21)`. The new cells copy column F's formatting, and column G has the same width as F. - **Summary:** A8 is "Total tax" and B8 is `=Sales!G22`. The two "TEST" rows (8 and 15) are included in the tax and in the totals, as they already are for Revenue and Commission.

files: `demo/runs/live-sonnet/t1_tax_column__default__3/` (book.xlsx, transcript)

### Add a computed column — Tabula (TEL) (trial 1): ✅ correct  ·  $0.033 · 11s · 4 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column and the Summary line, and saved them to `book.xlsx`. Tabula's check passed with no errors, and the apply step re-read the file and verified it. - **Sales sheet:** G1 is "Tax". G2:G21 is `=E2*TaxRate`, filled down, where `TaxRate` is the named cell `Inputs!B2`. G22 is `=SUM(G2:G21)`. - **Summary sheet:** A8 is "Total tax" and B8 is `=Sales!G22`. All the new cells are live formulas. G8 comes out at only 0.18, so that row's Revenue (E8) looks like about 1. The other rows are in the hundreds or thousands. I didn't change it, but you may want to check that row's source data.

files: `demo/runs/live-sonnet/t1_tax_column__tabula__1/` (book.xlsx, transcript)

### Add a computed column — Tabula (TEL) (trial 2): ✅ correct  ·  $0.032 · 10s · 4 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column and the Summary line, and saved them to `book.xlsx` in place. Tabula's check passed, and `apply` re-read the file and verified the 24 written cells. I haven't opened the file in Excel. - **Sales sheet:** G1 is "Tax". G2:G21 each use `=E2*TaxRate`, where `TaxRate` is the existing name for `Inputs!B2`. G22 is `=SUM(G2:G21)`. - **Summary sheet:** A8 is "Total tax" and B8 is `=Sales!G22`. G8 comes out as 0.18, which is very small next to the other rows (567 to 907). That looks like row 8's Revenue is unusually low, which I didn't investigate. I also left the script `edit.tel…

files: `demo/runs/live-sonnet/t1_tax_column__tabula__2/` (book.xlsx, transcript)

### Add a computed column — Tabula (TEL) (trial 3): ✅ correct  ·  $0.033 · 10s · 4 tool calls

- ✅ G1 header is 'Tax'
- ✅ tax per row = revenue x tax rate
- ✅ tax cells are formulas
- ✅ tax total in row 22
- ✅ Summary A8 label mentions tax
- ✅ Summary B8 = total tax
- ✅ Summary B8 is a formula
- ✅ existing Summary figures unchanged
- ✅ tax follows the tax rate (live, not hard-coded)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the Tax column to the Sales sheet and a "Total tax" line to the Summary sheet, and saved `book.xlsx` in place. Tabula's check passed, and the apply step re-read the file and confirmed the 24 written cells. I haven't opened the file in Excel. - **Sales:** G1 is "Tax". G2:G21 each hold `=E2*TaxRate` (and so on down), where `TaxRate` is the defined name for `Inputs!B2`. G22 is `=SUM(G2:G21)`. - **Summary:** A8 is "Total tax" and B8 is `=Sales!G22`. One value looks odd. G8 comes out at 0.18, which means row 8's Revenue is tiny. I left it alone because I only added the new column. I also le…

files: `demo/runs/live-sonnet/t1_tax_column__tabula__3/` (book.xlsx, transcript)

### Insert a missed sale — default (Python + openpyxl) (trial 1): ❌ SILENT FAILURE  ·  $0.034 · 10s · 2 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ❌ revenue/commission per row — Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)
- ✅ Total row is row 23
- ❌ Sales totals include exactly the data rows — Sales!E23=82133 (expected 81269); Sales!F23=4,959 (expected 4,876)
- ❌ every Summary figure correct — Summary!B3=82133 (expected 81269); Summary!B4=4,959 (expected 4,876); Summary!B5=3,911 (expected 3,870); Summary!B6=18271 (expected 21816)
- ❌ Summary follows a change to the first sale — Summary!B3=89093 (expected 88229)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the sale as row 11 on the Sales sheet: West, Dana Lee, 40 units at 125. I haven't opened the saved file in Excel or recalculated it, so the new values haven't been checked. - **New row:** Revenue is `=C11*D11` and Commission is `=E11*CommissionRate`. The formatting is copied from row 10. - **Totals:** The Total row moved to row 23. Its sums now cover rows 2–22. - **Summary sheet:** Total revenue and Total commission now point to `Sales!E23` and `Sales!F23`. North revenue sums rows 2–22. The named range `SalesRevenue`, which Average sale uses, now spans `E2:E22`. The Sales sheet has two…

files: `demo/runs/live-sonnet/t2_insert_row__default__1/` (book.xlsx, transcript)

### Insert a missed sale — default (Python + openpyxl) (trial 2): ✅ correct  ·  $0.053 · 28s · 5 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 23
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the West / Dana Lee sale as row 11 of the Sales sheet and saved it to `book.xlsx`. I haven't opened the file in Excel or recalculated it, so the formula results are unchecked. I printed the edited rows after saving, and their formulas look right. - **New row:** Units is 40 and Unit Price is 125. Revenue is `=C11*D11` and Commission is `=E11*CommissionRate`, the same as the other rows. I copied the formatting from row 10. - **Shifted rows:** The old rows 11–21 moved down to 12–22. I rewrote their Revenue and Commission formulas so each one points at its own row. - **Totals:** The Total …

files: `demo/runs/live-sonnet/t2_insert_row__default__2/` (book.xlsx, transcript)

### Insert a missed sale — default (Python + openpyxl) (trial 3): ❌ SILENT FAILURE  ·  $0.031 · 14s · 2 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ❌ revenue/commission per row — Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)
- ✅ Total row is row 23
- ❌ Sales totals include exactly the data rows — Sales!E23=82133 (expected 81269); Sales!F23=4,959 (expected 4,876)
- ❌ every Summary figure correct — Summary!B3=82133 (expected 81269); Summary!B4=4,959 (expected 4,876); Summary!B5=3,911 (expected 3,870); Summary!B6=18271 (expected 21816)
- ❌ Summary follows a change to the first sale — Summary!B3=89093 (expected 88229)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I added the West / Dana Lee sale (40 units at 125) as row 11 on the Sales sheet. I haven't opened the file in Excel to recalculate it, so the new totals haven't been checked. - **New row:** Revenue is `=C11*D11` and Commission is `=E11*CommissionRate`, the same as the other rows. It uses the same formatting as row 10. - **Sales totals:** The old rows 11–21 moved down one row, and the Total row is now row 23. Its sums now run over rows 2–22. - **Summary sheet:** - Total revenue and Total commission now point to `Sales!E23` and `Sales!F23`. - North revenue now sums rows 2–22. - Average sale uses…

files: `demo/runs/live-sonnet/t2_insert_row__default__3/` (book.xlsx, transcript)

### Insert a missed sale — Tabula (TEL) (trial 1): ✅ correct  ·  $0.040 · 14s · 6 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 23
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I inserted the new sale as row 11 on the Sales sheet and saved `book.xlsx` in place. Tabula re-read the saved file and verified it. I didn't open it in Excel. - **New row:** West, Dana Lee, 40 units at 125. Revenue is `=C11*D11`, which gives 5,000. Commission is `=E11*CommissionRate`, which gives 300. - **Totals:** The totals row moved to row 23. Its `SUM` ranges grew to cover the new row, so Units is 657, Revenue is 81,269 and Commission is 4,876.14. - **Summary sheet:** - Total revenue (B3) is 81,269. - Total commission (B4) is 4,876.14. - Average revenue (B5) is about 3,869.95. - The North-…

files: `demo/runs/live-sonnet/t2_insert_row__tabula__1/` (book.xlsx, transcript)

### Insert a missed sale — Tabula (TEL) (trial 2): ✅ correct  ·  $0.036 · 12s · 5 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 23
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I inserted the West / Dana Lee sale (40 units at 125) as row 11 on the Sales sheet and saved `book.xlsx` in place. The apply step re-read the file and verified it. - **New row:** Revenue is `=C11*D11`, which gives 5,000. Commission is `=E11*CommissionRate`, which gives 300. These match the other rows. - **Totals:** Inserting inside the range made every total formula and named range cover the new row. The Sales totals row is now row 23, and its totals are Units 657, Revenue 81,269 and Commission 4,876.14. - **Summary:** All five figures now include the new sale. - Total revenue (`B3`) is 81,269…

files: `demo/runs/live-sonnet/t2_insert_row__tabula__2/` (book.xlsx, transcript)

### Insert a missed sale — Tabula (TEL) (trial 3): ✅ correct  ·  $0.036 · 11s · 5 tool calls

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 23
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I inserted the missing sale as row 11 on the Sales sheet and saved `book.xlsx` in place. Tabula re-read the file after writing and verified it. - **New row:** West, Dana Lee, 40 units at 125. Revenue is `=C11*D11`, which gives 5,000. Commission is `=E11*CommissionRate`, which gives 300. - **Totals:** The total row moved to row 23, and its sums now cover the new row. - Units: 617 → 657 - Revenue: 76,269 → 81,269 - Commission: 4,576.14 → 4,876.14 - **Summary sheet:** Every figure now includes the sale. - Total revenue: 81,269 - Total commission: 4,876.14 - Average revenue (`SalesRevenue`): 3,813…

files: `demo/runs/live-sonnet/t2_insert_row__tabula__3/` (book.xlsx, transcript)

### Delete test rows — default (Python + openpyxl) (trial 1): ❌ SILENT FAILURE  ·  $0.033 · 10s · 2 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ❌ every Summary figure correct — Summary!B3=0 (expected 76267); Summary!B4=0 (expected 4,576); Summary!B7='No' (expected 'Yes')
- ❌ Summary follows a change to the first sale — the result no longer depends on Sales!C2 (the link is broken)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows (original rows 8 and 15) from the Sales sheet, and the rows below moved up. I haven't opened the saved file in Excel or recalculated it, so I haven't seen the new totals. The data now ends at row 19 and the Total row is at row 20. - **Row formulas:** Revenue and Commission in rows 2–19 still use live formulas that point at their own row. - **Totals:** The Total row now sums rows 2–19 for Units, Revenue and Commission. - **Summary:** The North revenue formula now covers Sales rows 2–19. Total revenue and Total commission point at the Total row's cells, so they follow…

files: `demo/runs/live-sonnet/t3_delete_test_rows__default__1/` (book.xlsx, transcript)

### Delete test rows — default (Python + openpyxl) (trial 2): ✅ correct  ·  $0.032 · 10s · 2 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows (old rows 8 and 15) from the Sales sheet, so the rows below moved up. I haven't opened the file in Excel to check the recalculated values. I only printed the saved formulas. - **Sales rows:** The Revenue and Commission formulas in each remaining row now point at their own row. The Total row is now row 20 and sums rows 2–19. - **Summary:** Total revenue and Total commission now point to `Sales!E20` and `Sales!F20`. The North revenue `SUMIF` covers rows 2–19. - **Named range:** I also changed `SalesRevenue`, which the Average sale formula uses, from `E2:E21` to `E2:E1…

files: `demo/runs/live-sonnet/t3_delete_test_rows__default__2/` (book.xlsx, transcript)

### Delete test rows — default (Python + openpyxl) (trial 3): ❌ SILENT FAILURE  ·  $0.030 · 10s · 2 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ❌ every Summary figure correct — Summary!B3=0 (expected 76267); Summary!B4=0 (expected 4,576); Summary!B7='No' (expected 'Yes')
- ❌ Summary follows a change to the first sale — the result no longer depends on Sales!C2 (the link is broken)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows (original rows 8 and 15) from the Sales sheet. The rows below moved up, so the data now runs from row 2 to row 19 and the Total row is row 20. I didn't recalculate the workbook or open it in Excel, so the formulas and totals are untested. A check after the save only printed the last Sales rows, and they look right. - **Row formulas:** the Revenue and Commission formulas in rows 2–19 are rewritten to point at their own rows. - **Totals:** the Total row now sums rows 2–19 as live `SUM` formulas. - **Named range:** `SalesRevenue` now covers `Sales!$E$2:$E$19`. The Aver…

files: `demo/runs/live-sonnet/t3_delete_test_rows__default__3/` (book.xlsx, transcript)

### Delete test rows — Tabula (TEL) (trial 1): ✅ correct  ·  $0.038 · 12s · 5 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows from the Sales sheet (original rows 8 and 15) and saved `book.xlsx` in place. The rows below moved up, so the Sales data now ends at row 19 and the totals are in row 20. - **Formulas:** Tabula's check passed, and I added a check that the Sales revenue total equals the sum of the remaining rows, which also passed. The totals and Summary figures still use live formulas. They now read `SUM(C2:C19)`, `SUM(E2:E19)` and `SUM(F2:F19)`. The Summary formulas and the `SalesRevenue` named range were shifted to match. - **Value changes:** The test rows had small values, so the …

files: `demo/runs/live-sonnet/t3_delete_test_rows__tabula__1/` (book.xlsx, transcript)

### Delete test rows — Tabula (TEL) (trial 2): ✅ correct  ·  $0.041 · 14s · 6 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows (original rows 8 and 15) from the Sales sheet and saved `book.xlsx` in place. - **Sales:** The rows below moved up. The data now covers rows 2–19, and the totals are in row 20: `=SUM(C2:C19)`, `=SUM(E2:E19)` and `=SUM(F2:F19)`. - **Summary:** All formulas moved with the rows and are still live. `B3` is `=Sales!E20` and `B4` is `=Sales!F20`. `B6` is `=SUMIF(Sales!A2:A19,"North",Sales!E2:E19)`. - **Named range:** `SalesRevenue` now covers `Sales!$E$2:$E$19`. - **Value changes:** The test rows' units and revenue were removed from the totals, so these figures shifted sl…

files: `demo/runs/live-sonnet/t3_delete_test_rows__tabula__2/` (book.xlsx, transcript)

### Delete test rows — Tabula (TEL) (trial 3): ✅ correct  ·  $0.039 · 12s · 5 tool calls

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I deleted the two TEST rows (rows 8 and 15) from the Sales sheet and saved the result to `book.xlsx`. The rows below moved up, and the Sales data now runs from row 2 to row 19. - **Formulas:** Tabula rewrote all 30 affected formulas to follow the moved rows, and they are all still live formulas. The Sales totals are now in row 20 (`=SUM(C2:C19)`, `=SUM(E2:E19)`, `=SUM(F2:F19)`). - **Summary sheet:** `B3` and `B4` now point at `Sales!E20` and `Sales!F20`. The `B6` `SUMIF` ranges and the `SalesRevenue` named range both shrank to rows 2–19. - **Figures:** Revenue went from 76,269 to 76,267 and co…

files: `demo/runs/live-sonnet/t3_delete_test_rows__tabula__3/` (book.xlsx, transcript)

### Fix two wrong figures — default (Python + openpyxl) (trial 1): ✅ correct  ·  $0.028 · 10s · 3 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong figures on the Summary sheet and saved `book.xlsx` in place. I haven't recalculated or opened the file, so the new values are unchecked. Both are still live formulas, and no other sheet was changed. - **Total commission (B4):** it pointed at total revenue (`=Sales!E22`). It now reads `=Sales!F22`, the commission total. - **North revenue (B6):** the `SUMIF` was matching "North" against the Rep column (B), so it would have returned 0. It now matches against the Region column: `=SUMIF(Sales!A2:A21,"North",Sales!E2:E21)`. Sales rows 8 and 15 are "Test" rows with 1 unit at $1.…

files: `demo/runs/live-sonnet/t4_fix_summary__default__1/` (book.xlsx, transcript)

### Fix two wrong figures — default (Python + openpyxl) (trial 2): ✅ correct  ·  $0.026 · 10s · 3 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong figures on the Summary sheet and saved `book.xlsx` in place. I haven't opened the file in Excel to check the new values. - **Total commission (B4):** it pointed at the total revenue cell (`Sales!E22`). It now reads `=Sales!F22`, the commission total. - **North revenue (B6):** the SUMIF looked for "North" in the Rep column (B), so it would have returned 0. It now looks in the Region column: `=SUMIF(Sales!A2:A21,"North",Sales!E2:E21)`. Both are still live formulas, and I didn't touch the other sheets. The Sales data has two "Test" rows (rows 8 and 15). They feed into total …

files: `demo/runs/live-sonnet/t4_fix_summary__default__2/` (book.xlsx, transcript)

### Fix two wrong figures — default (Python + openpyxl) (trial 3): ✅ correct  ·  $0.025 · 10s · 3 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong figures on the Summary sheet and saved `book.xlsx` in place. I haven't recalculated the workbook or checked the resulting values. - **Total commission (B4):** it pointed at the total revenue cell (`Sales!E22`). It now reads `=Sales!F22`, the commission total. - **North revenue (B6):** the SUMIF tested the Rep column (B), so it could never match "North". It now tests the Region column: `=SUMIF(Sales!A2:A21,"North",Sales!E2:E21)`. The other Summary figures and the other sheets are unchanged. The Sales data has two "Test" rows (rows 8 and 15) that are included in the totals …

files: `demo/runs/live-sonnet/t4_fix_summary__default__3/` (book.xlsx, transcript)

### Fix two wrong figures — Tabula (TEL) (trial 1): ✅ correct  ·  $0.037 · 12s · 5 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong figures on the Summary sheet and saved `book.xlsx` in place. Both are still live formulas, and nothing on the other sheets changed. Tabula checked the edit and re-read the file after writing. I haven't opened it in Excel. - **Total commission (`B4`):** It was `=Sales!E22`, which is total revenue. It is now `=Sales!F22`, the commission total, and shows 4576.14 (was 76269). - **North revenue (`B6`):** The `SUMIF` was matching "North" against the Rep column (`Sales!B`), so it returned 0. It now matches against the Region column (`Sales!A2:A21`) and shows 21816. The other Sum…

files: `demo/runs/live-sonnet/t4_fix_summary__tabula__1/` (book.xlsx, transcript)

### Fix two wrong figures — Tabula (TEL) (trial 2): ✅ correct  ·  $0.036 · 12s · 5 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong figures on the Summary sheet and saved `book.xlsx` in place. Both are still live formulas, and I didn't touch the other sheets. - **Total commission (`B4`):** it pointed at the revenue total (`Sales!E22`). It now reads `=Sales!F22` and shows 4576.14 instead of 76269. - **North revenue (`B6`):** the `SUMIF` tested the Rep column (`Sales!B`), so it matched nothing and returned 0. It now tests the Region column: `=SUMIF(Sales!A2:A21,"North",Sales!E2:E21)`. It shows 21816. I didn't change the other Summary figures (`B3`, `B5`, `B7`) because their formulas looked right. The va…

files: `demo/runs/live-sonnet/t4_fix_summary__tabula__2/` (book.xlsx, transcript)

### Fix two wrong figures — Tabula (TEL) (trial 3): ✅ correct  ·  $0.041 · 12s · 6 tool calls

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I fixed the two wrong Summary figures in `book.xlsx` and saved the file in place. Both are still live formulas, and the other sheets are untouched. - **Total commission (`Summary!B4`):** it was `=Sales!E22`, which is total revenue. It is now `=Sales!F22`, the commission total, and shows 4576.14. - **North revenue (`Summary!B6`):** the SUMIF tested the Rep column (`Sales!B2:B21`) for "North", so it returned 0. It now tests the Region column (`Sales!A2:A21`) and shows 21816. The check passed with no errors and `apply` re-read the saved file to verify it. The 4576.14 and 21816 are Tabula's predic…

files: `demo/runs/live-sonnet/t4_fix_summary__tabula__3/` (book.xlsx, transcript)

### Rename a sheet — default (Python + openpyxl) (trial 1): ✅ correct  ·  $0.020 · 9s · 2 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the sheet "Sales" to "Q3 Sales" in `book.xlsx` and saved it in place. I haven't opened the file in Excel to check the calculated values. - **Formulas:** I updated the three formulas on Summary that pointed at the old name, in B3, B4 and B6. They now use `'Q3 Sales'!`, and they are still live formulas. - **Named range:** `SalesRevenue` now points to `'Q3 Sales'!$E$2:$E$21`. The Summary formula that uses it, `AVERAGE(SalesRevenue)` in B5, needed no change. - **Other named ranges:** `TaxRate`, `CommissionRate` and `Target` point to the Inputs sheet, so I left them alone. I only searched…

files: `demo/runs/live-sonnet/t5_rename_sheet__default__1/` (book.xlsx, transcript)

### Rename a sheet — default (Python + openpyxl) (trial 2): ✅ correct  ·  $0.019 · 8s · 2 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the sheet "Sales" to "Q3 Sales" and saved the change to `book.xlsx` in place. I didn't recalculate the workbook or open it in Excel. I only listed the formulas afterwards to check the new references. - **Formulas:** Summary `B3`, `B4` and `B6` now point to `'Q3 Sales'!…`. They are still live formulas. - **Named range:** `SalesRevenue` now refers to `'Q3 Sales'!$E$2:$E$21`, so `AVERAGE(SalesRevenue)` in Summary `B5` still works. - **Other names:** The other named ranges point at the Inputs sheet, so I left them alone.

files: `demo/runs/live-sonnet/t5_rename_sheet__default__2/` (book.xlsx, transcript)

### Rename a sheet — default (Python + openpyxl) (trial 3): ✅ correct  ·  $0.019 · 9s · 2 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the "Sales" sheet to "Q3 Sales" and saved `book.xlsx` in place. I didn't reopen the saved file or recalculate it in Excel to check the values. - **Formulas on Summary:** B3, B4 and B6 now point to `'Q3 Sales'!`. They are still live formulas. For example, B6 is `=SUMIF('Q3 Sales'!A2:A21,"North",'Q3 Sales'!E2:E21)`. - **Named range:** `SalesRevenue` now refers to `'Q3 Sales'!$E$2:$E$21`. Summary B5, `=AVERAGE(SalesRevenue)`, didn't need editing. - **Other names:** `TaxRate`, `CommissionRate` and `Target` point to the Inputs sheet and were left alone. Before editing, I searched every sh…

files: `demo/runs/live-sonnet/t5_rename_sheet__default__3/` (book.xlsx, transcript)

### Rename a sheet — Tabula (TEL) (trial 1): ✅ correct  ·  $0.032 · 12s · 4 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the sheet "Sales" to "Q3 Sales" and saved `book.xlsx` in place. - **Formulas on Summary:** B3, B4 and B6 now point at `'Q3 Sales'`. B5 and B7 still use the names `SalesRevenue` and `Target`, so they needed no change. - **Named range:** `SalesRevenue` now points to `'Q3 Sales'!$E$2:$E$21`. The other names point at Inputs and are unchanged. - **Checks:** the tool reported 3 formulas rewritten, and my check that `Summary!B3` equals `'Q3 Sales'!E22` passed. I didn't open the file in Excel. The edit script is saved as `edit.tel` in the folder.

files: `demo/runs/live-sonnet/t5_rename_sheet__tabula__1/` (book.xlsx, transcript)

### Rename a sheet — Tabula (TEL) (trial 2): ✅ correct  ·  $0.031 · 10s · 4 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the sheet "Sales" to "Q3 Sales" and saved `book.xlsx` in place. - **Formulas:** the Summary formulas in B3, B4 and B6 now point at `'Q3 Sales'!…`. They are still live formulas. - **Named range:** `SalesRevenue` now points at `'Q3 Sales'!$E$2:$E$21`. The Summary formula `AVERAGE(SalesRevenue)` didn't need to change. - **Check:** I added an `expect` that `Summary!B3` equals `'Q3 Sales'!E22`, and it passed. Tabula also re-read the saved file and verified it. The edit script is `edit.tel`, left in the folder.

files: `demo/runs/live-sonnet/t5_rename_sheet__tabula__2/` (book.xlsx, transcript)

### Rename a sheet — Tabula (TEL) (trial 3): ✅ correct  ·  $0.029 · 8s · 4 tool calls

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: I renamed the sheet "Sales" to "Q3 Sales" and saved `book.xlsx` in place. Tabula applied the edit and re-read the file to verify it. - **Formulas:** The three formulas on Summary that pointed at Sales now point at 'Q3 Sales': B3, B4 and B6. They're still live formulas. - **Named range:** `SalesRevenue` now refers to `'Q3 Sales'!$E$2:$E$21`. Summary!B5, which uses that name, needed no change. - **Check:** An `expect` that Summary!B3 equals `'Q3 Sales'!E22` passed. The edit script is `edit.tel` in the same folder.

files: `demo/runs/live-sonnet/t5_rename_sheet__tabula__3/` (book.xlsx, transcript)

## Threats to validity

- The Tabula arm receives a language reference (TEL.md); the default arm relies on the model's existing knowledge of openpyxl.
- Small sample: one model, few trials per task; treat differences as illustrative, not statistically established.
- The checker covers the functions this workbook uses; pycel and Excel may differ on functions outside that set.
- Deterministic reference runs use hand-written naive openpyxl scripts; they show the failure mode, not how often a real agent hits it — the live runs measure that.
