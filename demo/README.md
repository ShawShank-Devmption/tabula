# A/B demo: an agent editing Excel with Tabula vs the default method

The same agent gets the same task on the same workbook, under two arms. An independent
checker then scores both outputs.

| | default arm | Tabula arm |
|---|---|---|
| How the agent edits | Python 3 + openpyxl (the usual route for agents editing `.xlsx`) | `./tabula` CLI + TEL scripts |
| Extra material | — | `TEL.md` (the language reference) |
| Everything else | identical: prompt, model, budget, clean Claude Code config (`--safe-mode`), same shell utilities | |

## Tasks (`tasks.py`)

| Task | What can go wrong |
|---|---|
| T1 Add a computed Tax column plus a summary line | baseline: no structural change |
| T2 Insert a missed sale after row 10 | `insert_rows` moves cells but rewrites no formulas or defined names |
| T3 Delete the two `TEST` rows | totals keep summing stale ranges, even themselves (a cycle) |
| T4 Fix two wrong Summary figures | debugging: needs inspection of formulas, not values |
| T5 Rename `Sales` to `Q3 Sales` | references break, and the new name needs quotes |

## Checks (`evaluate.py`)

- **Expected numbers** are computed in plain Python from `build_workbook.DATA`.
- **Cell values** in the agent's output are computed by **pycel**, an independent Excel formula
  engine. Tabula's engine is never used to judge Tabula.
- Each task checks:
  - the values are correct;
  - the cells hold formulas, not hard-coded numbers;
  - a liveness probe passes (change an input, and the result must follow);
  - there are no error values or circular references;
  - untouched data is intact.
- A **silent failure** is a run where the agent finished normally and reported success, but
  the workbook is wrong.

## Subscription, not API

- Live runs drive the `claude` CLI with your **claude.ai subscription login**.
- `run.py` refuses to start unless `claude auth status` reports a claude.ai login.
- It removes any `ANTHROPIC_API_KEY`-style variables from the agent's environment.
- The "cost" figures are Claude Code's estimate at API list price. On a subscription they count
  toward your plan's usage limits and are not billed.

## Running

```bash
python3 demo/build_workbook.py                     # (re)generate workbooks/
python3 -m venv demo/.venv && demo/.venv/bin/pip install pycel "openpyxl==3.0.10"
#   pycel needs openpyxl 3.0; the checker venv is separate from Tabula's own openpyxl 3.1

python3 demo/run.py --driver reference             # scripted reference solutions, no LLM
python3 demo/run.py --trials 3 --model sonnet      # live: headless Claude Code (claude -p)
python3 demo/run.py --tasks t2_insert_row --arms default,tabula --trials 5 --model haiku
demo/.venv/bin/python demo/evaluate.py demo/runs/<folder>   # re-score a folder
```

Each run folder `runs/<stamp>/<task>__<arm>__<trial>/` keeps:
- the agent's final `book.xlsx`;
- `transcript.jsonl` (every tool call);
- `prompt.txt`;
- `meta.json` (cost, time, tool calls).

`REPORT.md` and `results.json` are written next to the run folders.

`reference/` holds the deterministic solutions:
- TEL scripts for every task;
- straightforward openpyxl scripts. The T1 and T4 ones are correct. The T2, T3 and T5 ones
  do what openpyxl's API suggests, and that is what fails.

The reference runs show the failure *mode*. The live runs measure how often a real agent hits it.
