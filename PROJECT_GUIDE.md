# Project Guide — Tabula

This guide covers four things:
- **Part 0:** the Revision 2 extension in one page (use it for the faculty conversation);
- **Part 1:** a map of every file;
- **Part 2:** the Review 1 presentation script (kept as presented);
- **Part 3:** how to demo and defend Revision 2.

---

# Part 0 — The extension, in one minute

**One-line pitch:**
> "Agents now edit spreadsheets, and they do it blind. Tabula gives them a language whose
> compiler checks every edit against the real workbook before anything is written."

**What stayed (approved at Review 1):**
- the hand-written formula compiler (lexer, Pratt parser, evaluator);
- the dependency graph, cycle detection and incremental recomputation.

They are now the *verification engine*.

**What was added:**
- **TEL**, an edit language with its own front end (scoped symbol tables), an edit IR, dead-store
  elimination, simulation and a back end;
- reference relocation for structural edits (the linker relocation problem);
- `.xlsx` I/O with atomic, verified writes.

**Why it is worth it — the evidence is in the repo, not just the claim.**
`demo/` runs the same agent on the same five tasks with and without Tabula. Outputs are
scored by an independent formula engine. Three trials each:

| model | Tabula: tasks correct | default openpyxl: tasks correct | default: **silent failures** |
|---|---|---|---|
| Claude Sonnet 5.5 | 15/15 | 11/15 | 4 |
| Claude Haiku 4.5 | 15/15 | 9/15 | 6 |

A silent failure means the agent said it was done, but the workbook was wrong. Examples:
- row formulas off by one after an insert;
- Summary pointing at cells that a delete emptied;
- a named range left pointing at a renamed sheet.

With Haiku, Tabula was also cheaper ($0.044 vs $0.065 per task) and faster (37 s vs 51 s). The Tabula arm was re-run on the post-review build (tasks.md T3.11) with the same results.

See `demo/RESULTS.md`.

**What changed from Review 1 (be upfront; full table in requirements.md §15):**
- `^` is now **left-associative** (`=2^3^2` is 64, as in Excel; Review 1 said 512).
- `%` is postfix percent; `AVG` is `AVERAGE`.
- The bytecode VM moved to stretch. The edit IR, its optimiser and the back end cover the
  IR, optimisation and code-generation syllabus items instead.

---

# Part 1 — What every file does

## Documents

| File | What it is |
|---|---|
| `requirements.md` | Revision 2: problem, objectives, scope, 26 FRs + 8 NFRs, concept mapping, §15 changes from Review 1 |
| `design.md` | Revision 2 SRS: architecture, formula language (Excel subset), TEL grammar and semantics, IR, relocation rules, I/O safety, demo design, testing, trade-offs |
| `tasks.md` | Phase plan with live status marks |
| `docs/TEL.md` | TEL language reference (also handed to the demo agent) |
| `README.md` | Install, usage, layout |
| `demo/README.md` | How the A/B demo works and is run |
| `Phase1_Review1_Report_Tabula.pdf`, `Review1_Tabula_Presentation.pptx`, `docs_review1_source.html` | Review 1 deliverables, **left unchanged** as the record of what was presented |
| `archive/phase1-review1-snapshot.tar.gz` | The Phase 1 code and docs exactly as they were before Revision 2 |

## Implementation — `tabula/`

| File | Role | The one thing to remember |
|---|---|---|
| `refs.py` | References | Excel bounds (`XFD1048576`); `A:C` and `1:3` ranges; when a sheet name needs quotes |
| `values.py` | Value model | Excel coercions: `"2"+1` is 3; text compares case-insensitively; numbers < text < booleans |
| `lexer.py` | **Formula lexical analysis** | `Q3!A1` — a word before `!` is a sheet even if it looks like a cell |
| `parser.py` | **Formula syntax analysis** | Pratt loop, all binary operators left-associative (`lbp + 1`); postfix `%` below it |
| `emitter.py` | **Code generation** (formulas) | Minimal parentheses from binding powers; round-trip property is tested |
| `functions.py` | Function symbol table | 27 simulated vs ~400 known Excel names: tells "unsimulated" from "hallucinated" |
| `evaluator.py` | Interpretation | One formula at a time, never recursing into other cells |
| `relocate.py` | **Relocation** | Insert expands spanning ranges; delete shrinks them or yields `#REF!` |
| `workbook.py` | Global symbol table | Sheet-scoped names shadow workbook names |
| `graph.py` | Dependency analysis | Range edges bucketed by column; Kahn for order; iterative Tarjan for cycles |
| `engine.py` | Recompute + taint | Unknown values propagate as *unverified*, never guessed |
| `xlsx.py` | File boundary (only openpyxl user) | Replay structure → sync → temp → fidelity guard → verify → atomic rename |
| `tel/lexer.py`, `tel/parser.py` | **TEL front end** | Parser-controlled lexer modes (formula islands); newline-synchronised error recovery |
| `tel/resolver.py` | **Semantic analysis** | Scope chain, `let` inlining, shape checks, `E-NOSHEET`, suggestions |
| `tel/ir.py` | **IR + optimisation** | Dead-store elimination with read/structure barriers |
| `tel/simulate.py` | Analysis | Write-set, impact, new cycles, expects, all before writing |
| `tel/plan.py`, `tel/compiler.py`, `cli.py` | Driver and output | `check` never writes; `apply` writes only with zero errors |

## Tests — `tests/` (`python3 -m pytest tests/ -q` → 162 passed)

| File | Covers |
|---|---|
| `test_lexer.py`, `test_parser.py` | Token classes, Excel reference forms, precedence incl. left-assoc `^`, emitter round trip |
| `test_evaluator.py` | Excel semantics, aggregates, criteria, errors, cycles, a 3,000-deep chain |
| `test_gradebook.py` | The gradebook vs independently computed arithmetic |
| `test_relocate.py` | Excel's insert/delete/fill/rename behaviour table |
| `test_tel.py` | Every diagnostic code, scopes, dead writes, simulation, apply, atomicity, guards, fidelity |
| `test_regressions.py` | One test per bug found in the code review (see tasks.md T3.11) |

## Demo — `demo/`

| File | Purpose |
|---|---|
| `build_workbook.py` | Generates `workbooks/sales_q3.xlsx` (+ a buggy variant) from one `DATA` list |
| `tasks.py` | The 5 task prompts and expected figures (plain Python) |
| `run.py` | Runs both arms with headless Claude Code (or `--driver reference`) |
| `evaluate.py` | Independent checker (pycel) → `REPORT.md` |
| `reference/` | TEL solutions + naive openpyxl scripts |
| `runs/` | Recorded runs: `reference/`, `live-sonnet/`, `live-haiku/` |
| `RESULTS.md` | The headline comparison across runs |

`tools/make_gradebook.py` rebuilds the gradebook files. `tools/bench.py` runs the 10k-formula
benchmark (NFR-2).

---

# Part 2 — How the Review 1 deck was presented (kept as presented)

> Historical. Two statements below are superseded by Revision 2: `=2^3^2` is now 64
> (Excel is left-associative), and the evaluator no longer recurses (the deep-chain limit is fixed).

**Shape of it:** 19 slides, about 10–12 minutes, then viva. Slides 1–8 are the
case for the project; 9–14 are the design; 15–17 are proof it works; 18–19 close.
The examiners' marks sit mostly in slides 3, 5, 8, 9 and 15–17 — spend your time
there and move briskly through the rest.

**The sentence to open with, and keep returning to:**
> "Every cell formula is a small program, so a spreadsheet is really thousands of
> tiny programs plus a dependency graph — and building it properly needs a whole
> compiler."

### Slide-by-slide

| # | Slide | What to say (keep it to this) |
|---|---|---|
| 1 | Title | Name the project in one line: a spreadsheet engine where each formula goes through a full compiler pipeline. Don't read the title aloud. |
| 2 | Agenda | Five seconds. "Problem, design, prototype, roadmap." Move on. |
| 3 | Problem Statement | **Marks slide.** Make the two-layer split explicit: per-formula compilation, and per-sheet dependency analysis. Land the closing line: recomputing all N cells per edit is what makes a naive engine unusable. |
| 4 | Motivation | Why it's the right topic *for this course*: every classical phase is exercised, plus concepts a normal project never reaches. |
| 5 | Objectives | **Marks slide.** Don't read all ten. Say "front end, analysis, back end" and give one example from each column. |
| 6 | Scope | Point at the out-of-scope line deliberately — it shows you set boundaries rather than over-promised. |
| 7 | Background Study | Name your sources out loud: the Dragon Book, Pratt 1973, Sestoft on spreadsheet internals, Kahn 1962. Then the table: what each existing system leaves undone. |
| 8 | Compiler Concepts | **Highest-value slide.** Walk the left column top to bottom — this is the examiner's checklist that your project covers the syllabus. Pause on "cycle detection as a whole-program semantic check." |
| 9 | System Architecture | **Marks slide.** Trace one path with your finger: edit a cell → compile it → update the graph → check for cycles → schedule → evaluate. Then the dependency rule at the bottom. |
| 10 | Micro Layer | Follow `=SUM(A1:A3)*2` down the left column. Then the grey box: optimization removes the dependency on `B1`, so the optimizer changes the *graph*, not just the speed. That line is your innovation argument. |
| 11 | Macro Layer | Six steps, said quickly. Be explicit that the "1 edit → 14 cells" figure is a **Phase 3 target**, not a measurement — say the word "target". |
| 12 | Language Specification | Point to `power = unary [ "^" power ]` and say: that one recursive rule makes `^` right-associative, so `=2^3^2` is 512, not 64. It is the most concrete parsing fact you own. |
| 13 | Error Model | The idea in one sentence: errors are *values* that flow through formulas, not crashes. Hence one broken cell never takes the sheet down. |
| 14 | Innovation | Five claims, ten seconds each. Lead with incremental recomputation. |
| 15 | Prototype | Switch tone: "this is running today." 1,041 lines, seven modules, standard library only. Be upfront that the evaluator is interim and the VM replaces it in Phase 3. |
| 16 | Demonstration | **Do this live if you can** (`python3 -m tabula.cli examples/demo.txt`), with the slide as fallback. Narrate three beats: the edit propagates; the syntax error is located by caret; the cycle is caught — and everything else still evaluates. |
| 17 | Testing | 54 tests, 0.02 s. Call out two: `=2^3^2` proves associativity, and the untaken `1/0` branch proves `IF` is lazy. Mention the gradebook is cross-checked against a real Excel file. |
| 18 | Roadmap | Phase 1 done, Phase 2 analysis, Phase 3 optimizer and VM. Ten seconds. |
| 19 | Close | Restate the one-line pitch and stop talking. Let them ask. |

### Live-demo order (if you run it)

```
python3 -m pytest tests/ -q                     # 54 passed
python3 -m tabula.cli examples/gradebook.txt    # the realistic sheet
python3 -m tabula.cli examples/demo.txt         # errors and the cycle
```
Then open `examples/gradebook.xlsx` in Excel beside it and point out the numbers
match.

### Questions you should expect

| Question | Your answer |
|---|---|
| Why not use Lex/Yacc or ANTLR? | Hand-writing it means every algorithm is mine to explain, and I control the error messages — positions and carets came free. |
| How does Pratt parsing work? | Each operator has a binding power; the loop keeps consuming while the next operator binds tighter. Right-associativity is one flag: recurse at the same power instead of one higher. |
| How do you detect cycles? | Depth-first search with three-colouring — an edge back to a grey (in-progress) node is a cycle. In the prototype it's an active-cell set; Phase 2 makes it a proper graph walk that also reports the path. |
| Where is the intermediate code? | Phase 3: a stack bytecode with constant and reference pools. A formula is a single expression, so stack IR fits it better than three-address code. |
| What is actually optimized? | Constant folding, algebraic simplification, branch elimination, and a bytecode peephole — plus the big one, recomputing only dirty cells. |
| Isn't this just Excel? | Excel is the *problem statement*. The project is the machinery underneath, which Excel hides. |
| **What happens with a very deep chain?** | Be honest: the prototype evaluator recurses, so past 200 cells it hits Python's stack limit. `design.md` §11 already requires iterative DFS and Kahn for exactly this reason; Phase 2 fixes it. |

That last one is a known, documented limitation — say it before they find it.

---

# Regenerating things

```bash
python3 tools/make_gradebook.py     # rebuild both gradebook files
python3 -m pytest tests/ -q         # run all tests

# rebuild the PDF after editing docs_review1_source.html
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless \
  --disable-gpu --no-pdf-header-footer \
  --print-to-pdf="Phase1_Review1_Report_Tabula.pdf" docs_review1_source.html
```

---

# Part 3 — Demoing and defending Revision 2

### Live demo order

```
python3 -m pytest tests/ -q                                              # 162 passed
python3 -m tabula inspect demo/workbooks/sales_q3.xlsx                   # what an agent sees
python3 -m tabula check demo/workbooks/sales_q3.xlsx demo/reference/t2_insert_row.tel
#   -> writes, 6 affected formulas, 28 relocated, expects PASS
python3 -m tabula apply demo/workbooks/sales_q3.xlsx demo/reference/t2_insert_row.tel -o /tmp/out.xlsx
open /tmp/out.xlsx                                                       # Excel agrees
```

Then show a script full of mistakes (sheet typo, `TaxRat`, `SUMM`, range × number, list
shape): every error is caught in one run, each with a caret and a hint.

Finish on `demo/RESULTS.md`.

### Questions to expect

| Question | Answer |
|---|---|
| Why a *language*, not tool calls (MCP)? | A script is one transaction. The compiler sees all of it before running: scopes, dead writes, cycles and total impact. Tool calls are checked one at a time, after they have happened. |
| Isn't this Copilot / Claude in Excel? | Those are closed products whose plans are text for a human. Tabula's plan is *machine-checked* against the workbook. The closest research system (SheetMind) only validates grammar. |
| Why did `2^3^2` change? | Tabula now writes formulas that Excel evaluates. Excel's `^` is left-associative, so matching it is a correctness requirement, not a style choice. |
| Where are intermediate code and optimisation? | The edit IR (`tel/ir.py`): per-cell quadruple-like ops, with dead-store elimination using reads and structural edits as barriers. |
| Where is target code generation? | `emitter.py` (AST → Excel text) and the back end in `xlsx.py`, which replays structure and syncs contents into OOXML via openpyxl. |
| Why use openpyxl at all? | The file format is not a compiler concern. It is confined to `xlsx.py`; openpyxl evaluates nothing and relocates nothing, and Tabula does both. |
| How is insert-row relocation decided? | Like Excel: positions ≥ the insertion point move; a range that *spans* it grows; one that ends just above does not; deleted cells become `#REF!`. See `test_relocate.py`. |
| What can't Tabula simulate? | Functions outside the 27. They are parsed and tracked, and their values are marked *unverified* rather than guessed. |
| Is the demo fair? | Same prompt, model, budget and clean configuration. The checker is an independent engine (pycel). The threats to validity are listed in every report. |
| Why did the bytecode VM move to stretch? | The same concepts (IR, optimisation, code generation) are now exercised by the edit language, where they change outcomes. |
