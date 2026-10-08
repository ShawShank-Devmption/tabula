# Tabula vs default — A/B agent demo

- run folder: `reference`  ·  generated 2026-10-08 03:32
- driver: **reference**  (deterministic scripted solutions, no LLM)
- checker: independent formula engine (pycel) + plain-Python expected values; Tabula's engine is not used to judge Tabula

## Summary

| arm | tasks fully correct | checks passed | silent failures | avg cost (USD) | avg time (s) | avg tool calls |
|---|---|---|---|---|---|---|
| default (Python + openpyxl) | 2/5 | 34/47 | 3 | — | 0 | — |
| Tabula (TEL) | 5/5 | 47/47 | 0 | — | 0 | — |

*Silent failure*: the agent finished normally (no error, no timeout) but the workbook fails the checks — the failure an unchecked edit hides.

## Per task

| task | trap | default (Python + openpyxl) | Tabula (TEL) |
|---|---|---|---|
| Add a computed column | baseline: no structural change; both methods should manage | ✅ 11/11 | ✅ 11/11 |
| Insert a missed sale | openpyxl insert_rows moves cells but rewrites no formulas or defined names | ❌ 7/11 | ✅ 11/11 |
| Delete test rows | openpyxl delete_rows leaves totals summing stale ranges (even themselves: a cycle) | ❌ 4/9 | ✅ 9/9 |
| Fix two wrong figures | debugging: needs inspection of formulas, not just values | ✅ 9/9 | ✅ 9/9 |
| Rename a sheet | openpyxl renames the tab only; references to 'Sales' break, and the new name needs quotes | ❌ 3/7 | ✅ 7/7 |

## Runs

### Add a computed column — default (Python + openpyxl) (trial 1): ✅ correct  ·  0s

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

> agent's summary: added Tax column G and Summary!B8

files: `demo/runs/reference/t1_tax_column__default__1/` (book.xlsx, transcript)

### Add a computed column — Tabula (TEL) (trial 1): ✅ correct  ·  0s

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

> agent's summary: sha256: d806f82d88b78151e7ebb448686baea017fbd7f9ef2daf406ca93d856c7e3353

files: `demo/runs/reference/t1_tax_column__tabula__1/` (book.xlsx, transcript)

### Insert a missed sale — default (Python + openpyxl) (trial 1): ❌ SILENT FAILURE  ·  0s

- ✅ new sale in row 11
- ✅ row 11 revenue and commission
- ✅ row 11 calculations are formulas
- ✅ data rows intact and in order
- ❌ revenue/commission per row — Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)
- ✅ Total row is row 23
- ❌ Sales totals include exactly the data rows — Sales!C23=610 (expected 657); Sales!E23=77653 (expected 81269); Sales!F23=4,708 (expected 4,876)
- ❌ every Summary figure correct — Summary!B3=4480 (expected 81269); Summary!B4=250.9 (expected 4,876); Summary!B5=3,883 (expected 3,870); Summary!B6=13791 (expected 21816) (+1 more)
- ❌ Summary follows a change to the first sale — the result no longer depends on Sales!C2 (the link is broken)
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: inserted row 11

files: `demo/runs/reference/t2_insert_row__default__1/` (book.xlsx, transcript)

### Insert a missed sale — Tabula (TEL) (trial 1): ✅ correct  ·  0s

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

> agent's summary: sha256: d806f82d88b78151e7ebb448686baea017fbd7f9ef2daf406ca93d856c7e3353

files: `demo/runs/reference/t2_insert_row__tabula__1/` (book.xlsx, transcript)

### Delete test rows — default (Python + openpyxl) (trial 1): ❌ SILENT FAILURE  ·  0s

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ❌ revenue/commission per row — Sales!E8=4620 (expected 3520); Sales!E9=4340 (expected 4620); Sales!E10=4275 (expected 4340); Sales!E11=4940 (expected 4275) (+8 more)
- ✅ Total row is row 20
- ❌ Sales totals include exactly the data rows — Sales!C20='circular reference' (expected 615); Sales!E20='circular reference' (expected 76267); Sales!F20='circular reference' (expected 4,576)
- ❌ every Summary figure correct — Summary!B3='circular reference' (expected 76267); Summary!B4='circular reference' (expected 4,576); Summary!B5='circular reference' (expected 4,237); Summary!B6='circular reference' (expected 21816) (+1 more)
- ❌ Summary follows a change to the first sale — the result no longer depends on Sales!C2 (the link is broken)
- ❌ no error values or broken/circular formulas — Sales!F16: circular reference; Sales!E18: circular reference; Sales!F18: circular reference; Sales!C20: circular reference
- ✅ sheet Inputs unchanged

> agent's summary: deleted TEST rows

files: `demo/runs/reference/t3_delete_test_rows__default__1/` (book.xlsx, transcript)

### Delete test rows — Tabula (TEL) (trial 1): ✅ correct  ·  0s

- ✅ no TEST rows remain
- ✅ data rows intact and in order
- ✅ revenue/commission per row
- ✅ Total row is row 20
- ✅ Sales totals include exactly the data rows
- ✅ every Summary figure correct
- ✅ Summary follows a change to the first sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: sha256: d806f82d88b78151e7ebb448686baea017fbd7f9ef2daf406ca93d856c7e3353

files: `demo/runs/reference/t3_delete_test_rows__tabula__1/` (book.xlsx, transcript)

### Fix two wrong figures — default (Python + openpyxl) (trial 1): ✅ correct  ·  0s

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: fixed Summary!B4 and B6

files: `demo/runs/reference/t4_fix_summary__default__1/` (book.xlsx, transcript)

### Fix two wrong figures — Tabula (TEL) (trial 1): ✅ correct  ·  0s

- ✅ total commission fixed (B4)
- ✅ North revenue fixed (B6)
- ✅ fixes are formulas
- ✅ every Summary figure correct
- ✅ sheet Sales unchanged
- ✅ B4 follows the commission rate
- ✅ B6 follows a North sale
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: sha256: 67de1b00c761fa36d9a5b0faefe4f921c554a7dc2df8878547b4114bc25ba0ee

files: `demo/runs/reference/t4_fix_summary__tabula__1/` (book.xlsx, transcript)

### Rename a sheet — default (Python + openpyxl) (trial 1): ❌ SILENT FAILURE  ·  0s

- ✅ sheet renamed to 'Q3 Sales'
- ❌ every Summary figure correct — Summary!B3='refers to a sheet that does not exist' (expected 76269); Summary!B4='refers to a sheet that does not exist' (expected 4,576); Summary!B5='#NAME?' (expected 3,813); Summary!B6='refers to a sheet that does not exist' (expected 21816) (+1 more)
- ❌ named range SalesRevenue points at 'Q3 Sales' — SalesRevenue -> 'Sales!$E$2:$E$21'
- ✅ data rows intact and in order
- ❌ Summary follows the renamed sheet — refers to a sheet that does not exist
- ❌ no error values or broken/circular formulas — Summary!B3: refers to a sheet that does not exist; Summary!B4: refers to a sheet that does not exist; Summary!B5: #NAME?; Summary!B6: refers to a sheet that does not exist
- ✅ sheet Inputs unchanged

> agent's summary: renamed Sales to Q3 Sales

files: `demo/runs/reference/t5_rename_sheet__default__1/` (book.xlsx, transcript)

### Rename a sheet — Tabula (TEL) (trial 1): ✅ correct  ·  0s

- ✅ sheet renamed to 'Q3 Sales'
- ✅ every Summary figure correct
- ✅ named range SalesRevenue points at 'Q3 Sales'
- ✅ data rows intact and in order
- ✅ Summary follows the renamed sheet
- ✅ no error values or broken/circular formulas
- ✅ sheet Inputs unchanged

> agent's summary: sha256: d806f82d88b78151e7ebb448686baea017fbd7f9ef2daf406ca93d856c7e3353

files: `demo/runs/reference/t5_rename_sheet__tabula__1/` (book.xlsx, transcript)

## Threats to validity

- The Tabula arm receives a language reference (TEL.md); the default arm relies on the model's existing knowledge of openpyxl.
- Small sample: one model, few trials per task; treat differences as illustrative, not statistically established.
- The checker covers the functions this workbook uses; pycel and Excel may differ on functions outside that set.
- Deterministic reference runs use hand-written naive openpyxl scripts; they show the failure mode, not how often a real agent hits it — the live runs measure that.
