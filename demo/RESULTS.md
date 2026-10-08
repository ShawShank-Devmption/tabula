# A/B demo results: Tabula vs the default method

Each row is one run folder. The same agent, prompt, model and budget are used in both arms; outputs are scored by an independent formula engine (pycel) against plain-Python expected values. Regenerate with `demo/.venv/bin/python demo/evaluate.py --summary demo/RESULTS.md demo/runs/<folder>...`.

| runs | arm | tasks fully correct | silent failures | avg cost (USD) | avg time (s) |
|---|---|---|---|---|---|
| [reference](runs/reference/REPORT.md) — scripted reference (no LLM) | default (Python + openpyxl) | 2/5 | 3 | — | 0 |
| [reference](runs/reference/REPORT.md) — scripted reference (no LLM) | Tabula (TEL) | 5/5 | 0 | — | 0 |
| [live-sonnet](runs/live-sonnet/REPORT.md) — live agent: `claude-sonnet-5-5`, 3 trial(s) | default (Python + openpyxl) | 11/15 | 4 | 0.029 | 11 |
| [live-sonnet](runs/live-sonnet/REPORT.md) — live agent: `claude-sonnet-5-5`, 3 trial(s) | Tabula (TEL) | 15/15 | 0 | 0.036 | 11 |
| [live-haiku](runs/live-haiku/REPORT.md) — live agent: `claude-haiku-4-5-20251001`, 3 trial(s) | default (Python + openpyxl) | 9/15 | 6 | 0.065 | 51 |
| [live-haiku](runs/live-haiku/REPORT.md) — live agent: `claude-haiku-4-5-20251001`, 3 trial(s) | Tabula (TEL) | 15/15 | 0 | 0.044 | 37 |

**Silent failure**: the agent finished normally and reported success, but the workbook fails the checks.

## Per task (tasks fully correct / runs)

| task | trap | reference: default | reference: tabula | live-sonnet: default | live-sonnet: tabula | live-haiku: default | live-haiku: tabula |
|---|---|---|---|---|---|---|---|
| Add a computed column | baseline: no structural change; both methods should manage | 1/1 | 1/1 | 3/3 | 3/3 | 3/3 | 3/3 |
| Insert a missed sale | openpyxl insert_rows moves cells but rewrites no formulas or defined names | 0/1 | 1/1 | 1/3 | 3/3 | 1/3 | 3/3 |
| Delete test rows | openpyxl delete_rows leaves totals summing stale ranges (even themselves: a cycle) | 0/1 | 1/1 | 1/3 | 3/3 | 1/3 | 3/3 |
| Fix two wrong figures | debugging: needs inspection of formulas, not just values | 1/1 | 1/1 | 3/3 | 3/3 | 2/3 | 3/3 |
| Rename a sheet | openpyxl renames the tab only; references to 'Sales' break, and the new name needs quotes | 0/1 | 1/1 | 3/3 | 3/3 | 2/3 | 3/3 |

## What the silent failures looked like (live agents)

- **Insert a missed sale**, default (Python + openpyxl), `live-sonnet` trial 1: revenue/commission per row (Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)); Sales totals include exactly the data rows (Sales!E23=82133 (expected 81269); Sales!F23=4,959 (expected 4,876))
- **Insert a missed sale**, default (Python + openpyxl), `live-sonnet` trial 3: revenue/commission per row (Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)); Sales totals include exactly the data rows (Sales!E23=82133 (expected 81269); Sales!F23=4,959 (expected 4,876))
- **Delete test rows**, default (Python + openpyxl), `live-sonnet` trial 1: every Summary figure correct (Summary!B3=0 (expected 76267); Summary!B4=0 (expected 4,576); Summary!B7='No' (expected 'Yes')); Summary follows a change to the first sale (the result no longer depends on Sales!C2 (the link is broken))
- **Delete test rows**, default (Python + openpyxl), `live-sonnet` trial 3: every Summary figure correct (Summary!B3=0 (expected 76267); Summary!B4=0 (expected 4,576); Summary!B7='No' (expected 'Yes')); Summary follows a change to the first sale (the result no longer depends on Sales!C2 (the link is broken))
- **Insert a missed sale**, default (Python + openpyxl), `live-haiku` trial 2: revenue/commission per row (Sales!E12=5000 (expected 4340); Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370) (+7 more)); Sales totals include exactly the data rows (Sales!E23=82133 (expected 81269); Sales!F23=4,959 (expected 4,876))
- **Insert a missed sale**, default (Python + openpyxl), `live-haiku` trial 3: revenue/commission per row (Sales!E13=4340 (expected 4275); Sales!E14=4275 (expected 4940); Sales!E15=4940 (expected 4370); Sales!E16=4370 (expected 1) (+6 more)); Sales totals include exactly the data rows (Sales!E23=81473 (expected 81269); Sales!F23=4,880 (expected 4,876))
- **Delete test rows**, default (Python + openpyxl), `live-haiku` trial 2: every Summary figure correct (Summary!B5=8,028 (expected 4,237))
- **Delete test rows**, default (Python + openpyxl), `live-haiku` trial 3: revenue/commission per row (Sales!E8=4620 (expected 3520); Sales!E9=4340 (expected 4620); Sales!E10=4275 (expected 4340); Sales!E11=4940 (expected 4275) (+8 more)); Sales totals include exactly the data rows (Sales!E20=69143 (expected 76267); Sales!F20=7,769 (expected 4,576))
- **Fix two wrong figures**, default (Python + openpyxl), `live-haiku` trial 1: North revenue fixed (B6) (Summary!B6=0 (expected 21816)); every Summary figure correct (Summary!B6=0 (expected 21816))
- **Rename a sheet**, default (Python + openpyxl), `live-haiku` trial 1: every Summary figure correct (Summary!B5='#NAME?' (expected 3,813)); named range SalesRevenue points at 'Q3 Sales' (SalesRevenue -> 'Sales!$E$2:$E$21')
