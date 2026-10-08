# Tabula

**A statically checked spreadsheet edit language for AI agents, built on a compiled formula engine.**

Agents that edit Excel files usually write Python against openpyxl. openpyxl never evaluates
formulas and never rewrites references, so an inserted row silently breaks every total below
it, a renamed sheet breaks every formula that pointed at it, and nothing checks the edit
before it lands. Tabula gives agents a small language for edits, **TEL**. It compiles each
script against the real workbook and reports exactly what will change before writing a byte:

- references and names are resolved, and typos get suggestions;
- formulas are type-checked;
- the write-set and every affected formula are computed, with predicted values;
- circular references and failed `expect` checks are rejected.

Structural edits relocate every reference the way Excel does.

BCSE307P Compiler Design Laboratory — individual project. Design documents:
[`requirements.md`](requirements.md) (Revision 2), [`design.md`](design.md),
[`tasks.md`](tasks.md). Language reference: [`docs/TEL.md`](docs/TEL.md).

## Quick look

```
$ python3 -m tabula check demo/workbooks/sales_q3.xlsx demo/reference/t2_insert_row.tel
OK: no errors. Nothing has been written yet (this was a check).

writes (6)
  Sales!A11        West
  ...
  Sales!E11        =C11*D11  -> 5000
affected formulas (value changes) (6)
  Sales!E23        76269 -> 81269
  Summary!B3       76269 -> 81269
  Summary!B5       3813.45 -> 3869.95238095238
relocated: 28 formula(s) rewritten to follow moved/renamed cells
expects
  line 8    Sales!E23 = SUM(Sales!E2:E22)            PASS
```

A script with mistakes is rejected with positioned diagnostics:

```
error[E-SHEET] edit.tel:2:5: unknown sheet 'Summry'
   2 | set Summry!B8 =Sales!G22
     |     ^^^^^^
  hint: did you mean 'Summary'?
```

## Requirements

- Python 3.11+
- `openpyxl` 3.1 (`pip install -r requirements.txt`). It is used only at the file boundary
  (`tabula/xlsx.py`); the compilers, graph and engine are hand-written and use only the
  standard library.
- `pytest` for the test suite.
- The demo and the workspaces use the [`claude`](https://claude.com/claude-code) CLI, signed in
  with your Claude subscription (`claude auth login`). The demo also needs a small virtualenv for
  the independent checker (see [`demo/README.md`](demo/README.md)).

## Usage

```
python3 -m tabula inspect BOOK.xlsx [RANGE] [--json]    # sheets, headers, formula blocks, names
python3 -m tabula deps    BOOK.xlsx Sheet!A1 [--json]   # precedents / dependents
python3 -m tabula explain BOOK.xlsx Sheet!A1            # formula compiler pipeline for one cell
python3 -m tabula check   BOOK.xlsx EDIT.tel [--json]   # compile + simulate; never writes
python3 -m tabula apply   BOOK.xlsx EDIT.tel (-o OUT.xlsx | --in-place) [--json]
                          [--allow SPEC]... [--if-unchanged SHA256] [--allow-lossy]
python3 -m tabula repl    [SCRIPT.txt]                  # Phase 1 single-sheet REPL
python3 -m pytest tests/ -q                             # test suite
```

Exit codes: 0 OK · 1 the script has errors or apply was refused (nothing written) · 3 usage
or I/O error.

## TEL in one screen

```
in "Sales" {
  set G1 "Tax"
  set G2:G21 =E2*TaxRate          # fill: written for G2, copied down like Excel
  set G22 =SUM(G2:G21)
  insert rows 11                  # every reference on every sheet is relocated
  set A11:D11 ["West", "Dana Lee", 40, 125]
}
rename sheet "Inputs" to "Assumptions"
expect Summary!B3 = Sales!E23     # checked on the simulated result; failure rejects the script
```

Full reference, rules and error codes: [`docs/TEL.md`](docs/TEL.md).

## Try it with Claude Code (your Claude subscription)

```
python3 tools/make_workspace.py                    # demo workbook -> playground/tabula
python3 tools/make_workspace.py ~/Desktop/mine.xlsx --dir playground/mine   # your own workbook
python3 tools/make_workspace.py --method openpyxl  # the same workbook, plain Python (to compare)
cd playground/tabula && claude
```

What happens:
- Claude Code uses your claude.ai login (`claude auth status`). Nothing goes through an API key.
- On first start, accept the "trust this folder" prompt so the pre-approved commands apply.
- Then ask for edits in plain English. `PROMPTS.md` has the demo tasks.
- The workspace's `CLAUDE.md` tells Claude to inspect, write TEL, `check`, and then `apply`.
- Your original file is never modified: a copy is edited, and a pristine copy sits in `original/`.

## The A/B demo

`demo/` runs the same editing tasks on the same workbook twice: an agent (headless Claude
Code) editing with **Python + openpyxl**, and the same agent editing with **Tabula**. An
independent checker (pycel plus plain-Python expected values) scores both outputs.

```
python3 demo/run.py --driver reference    # deterministic, no LLM
python3 demo/run.py --trials 3            # live agents; writes demo/runs/<stamp>/REPORT.md
```

See [`demo/README.md`](demo/README.md).

## Layout

```
tabula/refs.py         A1 references, Excel bounds, sheet-name quoting
tabula/values.py       value model, Excel coercions and comparison
tabula/lexer.py        formula scanner
tabula/parser.py       formula Pratt parser + AST
tabula/emitter.py      AST -> Excel formula text
tabula/functions.py    simulated function registry + Excel function catalogue
tabula/evaluator.py    one-formula evaluator (Excel semantics)
tabula/relocate.py     fill shifting, insert/delete relocation, sheet rename
tabula/workbook.py     workbook model and symbol table
tabula/graph.py        dependency graph: Kahn, Tarjan, closure, cycle paths (all iterative)
tabula/engine.py       incremental recompute + taint; REPL facade
tabula/xlsx.py         openpyxl boundary: load, write back, fidelity guard, verification
tabula/diagnostics.py  coded, positioned diagnostics
tabula/tel/            TEL compiler: lexer, parser, resolver, IR, simulate, plan, compiler
tabula/cli.py          command line
docs/TEL.md            language reference
demo/                  A/B agent demo
tests/                 pytest suite
```
