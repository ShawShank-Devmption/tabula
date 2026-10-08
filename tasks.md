# Task Plan — Phase-Wise Implementation (Revision 2)

**Project:** Tabula — a statically checked spreadsheet edit language for AI agents, built on a compiled formula engine.
The phases follow the course: **Phase 1** design and prototype (Review 1, 20 marks) · **Phase 2** core implementation (Review 2, 25 marks) · **Phase 3** final implementation and testing (Review 3, 35 marks).

Each task lists a **Done when** criterion. Status as of 2026-10-08:
- `[x]` done and tested;
- `[~]` partly done;
- `[ ]` open.

Review dates follow the course schedule (not fixed here).

---

## Phase 1 — Problem Definition, Design, Prototype (done, Review 1)

- [x] Revision 1 requirements, design and task plan; Review 1 report (`Phase1_Review1_Report_Tabula.pdf`) and deck.
- [x] Formula lexer, Pratt parser, direct evaluator, REPL with `explain`; 54 tests; gradebook cross-checked against an Excel workbook.
- [ ] **T1.1 Git repository**: carried into Phase 2 (T2.0).

## Phase 2 — Core Implementation (Review 2)

**Goal:** demonstrate an agent's edit being compiled, checked and simulated against a real workbook, then applied safely. The formula compiler and the sheet engine become Excel-compatible and multi-sheet.

### 2.0 Housekeeping
- [ ] **T2.0** `git init`; first commit is the Phase 1 snapshot (`archive/phase1-review1-snapshot.tar.gz`); then commit per task.
- [ ] **T2.0a** **Faculty approval of the scope extension** (manual §4). Use requirements.md §15 and PROJECT_GUIDE "Part 0" for the conversation.

### 2.1 Excel-compatible formula compiler
- [x] **T2.1** References up to `XFD1048576`; sheet qualifiers (quoted and unquoted); `A:C`, `1:3`; error literals; postfix `%`; left-associative `^`; `_xlfn.` prefixes.
  *Done when:* lexer and parser tests cover each form → `tests/test_lexer.py`, `tests/test_parser.py`.
- [x] **T2.2** Emitter (AST → Excel text) with minimal parentheses and sheet quoting.
  *Done when:* round-trip property `parse(emit(parse(s))) == parse(s)` holds on a corpus.
- [x] **T2.3** Function registry (27 simulated functions) + Excel function catalogue; evaluator semantics (numeric text, case-insensitive comparison, cross-type order, aggregates skipping text, `ROUND` half away from zero, `#NUM!`, `COUNTIF`/`SUMIF` criteria with wildcards).
  *Done when:* `tests/test_evaluator.py` passes, including the Excel cases (`2^3^2` → 64, `"2"+1` → 3, `ROUND(2.5,0)` → 3).

### 2.2 Workbook engine
- [x] **T2.4** Workbook model and symbol table: sheets, defined names (sheet scope → workbook scope).
- [x] **T2.5** Dependency graph with cell edges and **symbolic range edges**; closure; Kahn; iterative Tarjan for cycles; cycle paths.
  *Done when:* a 3,000-deep chain evaluates without recursion (`test_deep_chain_has_no_recursion_limit`).
- [x] **T2.6** Incremental recompute; unverified-value (taint) propagation for unsimulated functions; recompute statistics.

### 2.3 `.xlsx` boundary
- [x] **T2.7** Load (formulas + Excel-cached values, dates → serials, unparsed → opaque); defined names; per-sheet structural blockers.
- [x] **T2.8** Write back: replay structure → sync contents → temp file → fidelity guard → verify by reloading → atomic rename. Lock-file and `--if-unchanged` guards.
  *Done when:* atomicity, guard and literal-text tests pass (`tests/test_tel.py`).

### 2.4 TEL compiler
- [x] **T2.9** Line scanner with parser-controlled formula mode; statement parser with newline-synchronised recovery.
- [x] **T2.10** Resolver: scope chain, `let`, sheet and name resolution with suggestions, `E-NOSHEET` rule, shape and size checks, function and arity checks, range-in-scalar-context, `W-SCOPE`.
- [x] **T2.11** Lowering to the edit IR, with fill expansion through `shift_for_fill`.
- [x] **T2.12** Simulation: write-set, `--allow`, hard-coding lint, new-cycle rejection, impact, new errors, unverified cells, `expect`.
- [x] **T2.13** Plan rendering (text + JSON schema 1) and the CLI (`inspect` with formula-block compression, `deps`, `explain`, `check`, `apply`, `repl`).

### 2.5 Demo and Review 2
- [x] **T2.14** A/B demo harness: generated workbooks, 5 tasks, pycel-based independent checker, runner for headless Claude Code (both arms isolated with `--safe-mode`), deterministic reference solutions, `REPORT.md`.
- [x] **T2.15** Live A/B runs recorded: `demo/runs/live-sonnet`, `demo/runs/live-haiku` (3 trials × 5 tasks × 2 arms each) plus `demo/runs/reference`. Headline in `demo/RESULTS.md`:
  - Tabula: 30/30 correct;
  - default openpyxl: 20/30, with 10 silent failures.

  The Tabula arm was re-run on the post-review build (T3.11) with the same 30/30 result. The default arm does not execute Tabula code, so its results are unaffected.
- [ ] **T2.16** Implementation progress report (module-wise explanation, screenshots of `check`/`apply`/`inspect`, demo results) per manual §10.3.
- [ ] **T2.17** Review 2 dry run, covering:
  - the live demo: `inspect` → a bad script (diagnostics) → a fixed script → `check` → `apply` → open in Excel;
  - the A/B report;
  - viva prep: scopes and symbol tables, relocation rules, dead-store elimination, Kahn and Tarjan, taint.

## Phase 3 — Completion, Evaluation, Documentation (Review 3)

### 3.1 Structural edits and relocation (built early because the demo needed them)
- [x] **T3.1** Relocation for insert/delete rows and columns (shift, expand, shrink, `#REF!`) across formulas, defined names and pending `expect`s; sheet rename with quoting.
  *Done when:* `tests/test_relocate.py` (Excel behaviour table) passes.
- [x] **T3.2** TEL structural statements, `add sheet`, `rename sheet`; refusal on merged cells / tables / CF / DV / unparsed formulas (`E-STRUCT`).
- [x] **T3.3** IR optimiser: dead-write elimination with `W-DEAD-WRITE`.

### 3.2 Evaluation
- [~] **T3.4** **Differential corpus.** Open the demo workbooks and the gradebook in Excel and save them, so they carry Excel-computed cached values; add three or more real-world workbooks. `inspect` then reports engine agreement. Add a pytest that asserts 100 % agreement on simulable formulas.
  *Now:* agreement reporting is implemented; pycel independently agrees with Tabula on the demo workbook. Excel-saved files are still needed.
- [x] **T3.5** Benchmark (`tools/bench.py`, 10,001 formulas: per-row, a 2,500-deep chain, windowed ranges, conditionals). Measured:
  - full recompute: 50 ms;
  - bounded edit: 20 formulas recomputed, simulation 15 ms;
  - edit at the top of the chain: 2,522 recomputed in 10 ms (recompute) / 94 ms (simulation);
  - `check` end to end: 0.4 s.

  NFR-2 is met. Column-bucketing the range edges fixed an O(V·R) closure: the chain edit went from 962 ms to 10 ms.
- [x] **T3.6** Fidelity-guard test: an injected `customXml` part produces `W-LOSSY` at check and `E-FIDELITY` at apply, leaves the input untouched, and `--allow-lossy` overrides (`test_fidelity_guard_refuses_lossy_save`).
- [~] **T3.7** A/B evaluation at scale: more trials, a second model, plus a task with merged cells (refusal path). Interpret the results in the report.

### 3.3 Documentation and Review 3
- [ ] **T3.8** Final report per manual §24 (reuse requirements/design; add test, benchmark and A/B results).
- [x] **T3.9** User instructions: `README.md`, `docs/TEL.md`, `demo/README.md`.
- [ ] **T3.10** Presentation and Review 3 rehearsal; evidence folder (terminal captures, REPORT.md, before/after workbooks opened in Excel).

### 3.4 Code review (2026-10-08)
- [x] **T3.11** Independent review of relocation, simulation, write-back, resolver and evaluator. 11 bugs plus 3 minor issues were found and fixed, each with a regression test in `tests/test_regressions.py`:
  - stale values behind whole-column ranges and relocated names after insert/delete;
  - `let` drifting under fill;
  - defined-name targets ignoring earlier inserts;
  - a crash with pre-existing cycles;
  - single-cell references and blanks in aggregates/criteria;
  - case-only rename;
  - `"#N/A"` text;
  - overflow;
  - references outside formulas (validation rules on other sheets, hyperlinks, print areas, hidden rows);
  - the `_xlfn.` storage prefix;
  - empty-string literals, `0^0`, and 15-digit comparison.

  The reviewer's 400-script randomized differential test (predicted state vs reloaded `apply` output) passes with 0 mismatches.

### Stretch (only if ahead)
- [ ] S1 Formula bytecode VM + AST optimiser (the Revision 1 Phase 3 plan), behind the same evaluator interface.
- [ ] S2 MCP server wrapper around the CLI.
- [ ] S3 Office Scripts code-generation target (run TEL inside Excel on the web).
- [ ] S4 Relocation of merged cells, conditional formats, data validations and tables (lifts most `E-STRUCT` refusals).

---

## Dependency Overview

```
T2.1 ─► T2.2 ─► T2.3 ─► T2.4 ─► T2.5 ─► T2.6 ─► T2.7 ─► T2.8
                                   │
T2.9 ─► T2.10 ─► T2.11 ─► T2.12 ─► T2.13 ─► T2.14 ─► T2.15 ─► T2.16/T2.17
                           │
                           └─► T3.1 ─► T3.2 ─► T3.3          (done)
T3.4 · T3.5 · T3.6 · T3.7 ─► T3.8 ─► T3.10
```

## Risk Register

| Risk | Impact | Mitigation |
|---|---|---|
| Faculty does not accept the extension | Rework of Review 2 framing | Extension keeps every Revision 1 component; requirements §15 lists each change with its reason; fall back to presenting TEL as the engine's "edit interface" |
| Excel semantics diverge from Tabula's predictions | Wrong predicted values | Differential testing against cached values (T3.4); unsupported features are opaque/unverified, never guessed |
| openpyxl drops workbook parts on save | Data loss for the user | Fidelity guard compares zip parts before replacing; post-write verification; atomic rename |
| Relocation bug corrupts formulas | Silent wrong workbook | Excel-behaviour table tests; refusal on features not relocated (`E-STRUCT`); `expect` lets agents state intent |
| A/B results are noisy or favour the default arm | Weaker demo claim | Report honestly with trials; the value of `check` (diagnostics, predicted impact) stands on its own; reference runs show the failure mode deterministically |
| Bytecode VM promised at Review 1 is now stretch | Viva question | Explain the trade (requirements §15): IR + optimiser + back end moved into the edit language |
| Review viva depth | Marks | Every module is hand-written and documented; PROJECT_GUIDE lists expected questions |
