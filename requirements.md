# Requirements Document

**Project Title:** Tabula — A Statically Checked Spreadsheet Edit Language for AI Agents, built on a Compiled Formula Engine with Dependency-Graph Analysis and Incremental Recomputation

**Course:** BCSE307P — Compiler Design Laboratory
**Project Type:** Individual
**Semester:** Fall 2026–27

> **Revision 2 (2026-10-08) — scope extension after Review 1.** Review 1 presented Tabula as a
> spreadsheet engine whose formulas are compiled. Revision 2 keeps that engine and puts a
> second language on top of it: **TEL (Tabula Edit Language)**, which AI agents use to read
> **and edit** real Excel workbooks. Every edit is compiled and checked against the actual
> workbook before anything is written. The formula compiler, the dependency graph and
> incremental recomputation are all retained; they become the verification engine for agent
> edits. A summary of changes from Revision 1 is in §15.

---

## 1. Project Overview

Tabula is a language-processing system that sits between an AI agent and an Excel workbook (`.xlsx`). It has two languages and a shared engine:

1. **The formula language** is an Excel-compatible subset (`=SUM(Sales!E2:E21)*TaxRate`). Each formula goes through a hand-written compiler pipeline: lexer, Pratt parser, semantic analysis, emitter and evaluator.
2. **TEL, the edit language**, is a small imperative language that agents write to change a workbook:
   ```
   in "Sales" {
     insert rows 11
     set A11:D11 ["West", "Dana Lee", 40, 125]
     set E11:F11 =C11*D11
   }
   expect Summary!B3 = Sales!E23
   ```
   The TEL compiler lexes, parses, resolves names against the workbook (sheets, defined names, `let` scopes) and type- and shape-checks the script. It then lowers the script to an edit IR, optimises that IR, analyses it, simulates it on the engine, and only after all of that writes the file.
3. **The workbook engine** is a multi-sheet cell store with a dependency graph, cycle detection and incremental recomputation in topological order. It predicts the value of every cell an edit touches, without needing Excel.
4. **An agent-facing CLI** (`python3 -m tabula inspect | deps | explain | check | apply | repl`) prints human-readable or JSON output, so any agent with a shell can use it.
5. **An A/B agent demo** runs the same editing tasks on the same workbook twice: once with Tabula, once with the default method (an agent writing Python/openpyxl). An independent checker then scores both outcomes.

The compilers, graph algorithms and engine are hand-written in Python. **openpyxl** is used only as the file-format layer for reading and writing `.xlsx`, as the manual (§18) permits for acknowledged libraries. It does not evaluate formulas or rewrite references, so all of that logic is Tabula's own.

---

## 2. Problem Statement

AI agents are increasingly asked to edit spreadsheets, not just read them. Today they do this either through product add-ins (Copilot Agent Mode, Claude in Excel) or, much more often, by writing Python against openpyxl. The second route has a structural blind spot:

- **openpyxl never evaluates formulas.** An agent that writes `=SUM(E2:E21)` cannot see the result until Excel opens the file.
- **openpyxl does not maintain references.** Its documentation says it "does not manage dependencies, such as formulae, tables, charts, etc., when rows or columns are inserted or deleted." So inserting one row silently breaks every total and cross-sheet reference below it. Deleting rows can even turn a `SUM` into a circular reference.
- **Nothing checks the edit before it is written.** Hallucinated sheet names, wrong target cells, values hard-coded where formulas should be, and newly created circular references all reach the file unnoticed.

Benchmarks confirm the gap. On SpreadsheetBench 2 (2026) the best model completes 34.9 % of real workflows, and the authors identify *insufficient inspection and incorrect target-cell selection* as the dominant failure causes. Prior agent systems constrain *syntax* at most: SheetMind uses a BNF grammar but validates nothing against the workbook. Commercial systems show a plan for a human to read, which nothing machine-checks.

**The problem:** give agents a way to express spreadsheet edits that a compiler can check *statically, against the real workbook*, before any byte is written. That means resolving every reference, computing exactly what will change, detecting cycles, maintaining references through structural edits, and predicting resulting values.

---

## 3. Motivation

- **It is a compiler problem twice over.** TEL has the full front end (lexing, parsing, scoped symbol tables, type and shape checking), an IR with an optimisation pass, and a code-generating back end. The formula language is a second, embedded language: the TEL lexer hands formula text to the formula compiler, a classic lexer-mode / island-grammar arrangement.
- **The Phase 1 engine gains a real job.** Dependency analysis answers "what does this edit affect?" (blast radius). Cycle detection rejects edits that would create circular references. Incremental recomputation predicts values quickly even on large workbooks.
- **Reference relocation is linker relocation.** Inserting rows must rewrite every address at or below the insertion point, across all sheets and defined names. Deleting rows must turn references to removed cells into `#REF!`. This is the relocation problem from linkers, applied to formulas.
- **The demo answers itself.** The same agent, the same task and the same workbook, edited two ways, then scored by an independent checker.

---

## 4. Objectives

1. Make the formula language **Excel-compatible** for its supported subset. That covers syntax (sheet-qualified references, whole-row and whole-column ranges, postfix `%`, error literals), precedence and associativity (`^` is left-associative in Excel), coercions and function semantics. Compatibility is checked against values cached by Excel.
2. Implement a **formula emitter** (AST → Excel formula text) with the round-trip property `parse(emit(ast)) == ast`.
3. Build the **workbook engine**. It holds multiple sheets and defined names (workbook- and sheet-scoped). Its dependency graph has cell edges plus symbolic range edges. It detects cycles with the path, and recomputes incrementally in topological order. Formulas it cannot simulate are treated as opaque and their effects are propagated as *unverified* (taint).
4. Implement **.xlsx loading and saving** through openpyxl. Loading reads formulas plus Excel-cached values. Saving is atomic, includes a fidelity guard that refuses writes that would lose workbook parts, and verifies the written file afterwards.
5. Design and implement **TEL**:
   - **lexer** with a formula mode;
   - **parser**;
   - **semantic analyser** (scope chain, declarations, sheet and name resolution with suggestions, shape and size checks);
   - **lowering to an edit IR**, including relative-reference fill expansion;
   - **dead-write elimination** on the IR;
   - **static analyses** (write-set, permissions, formula-overwrite lint, cycle check, blast radius);
   - **simulation** with `expect` assertions.
6. Implement **structural edits with reference relocation**: insert and delete rows and columns, add and rename sheets. Relocation covers all formulas, defined names and pending `expect`s.
7. Provide the **agent-facing CLI**: `inspect`, `deps`, `explain`, `check` (dry run, never writes), `apply` and `repl`. It supports a versioned JSON output schema and bounded output size.
8. Build the **A/B agent demo and evaluation**:
   - a generated demo workbook;
   - editing tasks with independent, value-based checkers;
   - an agent runner that drives a real agent through both methods;
   - a comparison report;
   - deterministic reference solutions, so the comparison can also be shown without an LLM.
9. Validate everything with a pytest suite: unit, scenario, golden output, differential (vs. Excel-cached values), round-trip and fidelity, and benchmark tests.

---

## 5. Scope

### In scope

| Area | Contents |
|---|---|
| Formula language | Excel syntax subset: number/string/boolean/error literals; references `A1`, `$A$1`, `Sheet!A1`, `'Q3 Sales'!A1:B4`, `A:A`, `1:3`; defined names; operators `+ - * / ^` (left-assoc), postfix `%`, `&`, comparisons, unary `-`; function calls |
| Supported functions (simulated) | `SUM AVERAGE MIN MAX COUNT COUNTA COUNTIF SUMIF ROUND ABS SQRT MOD POWER IF IFERROR AND OR NOT LEN CONCAT CONCATENATE UPPER LOWER TRIM ISNUMBER ISBLANK ISERROR` |
| Other Excel functions | Recognised from a catalogue of Excel function names: parsed, dependency-tracked, never simulated; their values come from Excel's cache and are marked *unverified* when inputs change |
| Workbook engine | Multi-sheet sparse cell store; defined names (workbook + sheet scope); dependency graph (cell + range edges); cycle detection with path; full + incremental recomputation; unverified-value propagation |
| TEL | `in` blocks, `let`, `set` (literal / list / formula with fill), `clear`, `expect`, `insert`/`delete` `rows`/`cols`, `add sheet`, `rename sheet`; comments |
| TEL compiler | Lexer with formula mode; parser; resolver with scope chain and symbol tables; edit IR; dead-write elimination; analyses; simulation; plan renderer (text + JSON); back end (replay + content sync to openpyxl) |
| Safety | Never write on any error; atomic save; post-write verification; fidelity guard; Excel lock-file detection; `--if-unchanged <sha256>` optimistic concurrency; `--allow` write allowlist |
| Tooling | CLI subcommands; formula `explain`; interactive REPL (Phase 1 commands kept) |
| Demo | Generated workbook, ≥ 4 tasks with checkers, agent runner (headless Claude Code, isolated), deterministic reference solutions, Markdown comparison report |
| Quality | pytest suite; differential tests against Excel-cached values; benchmark generator |

### Out of scope (future enhancements)

- Live Excel / Google Sheets (Office.js, Apps Script). Tabula edits `.xlsx` files on disk.
- Formatting and style edits, charts, pivot tables, images. These are preserved when the fidelity guard confirms it, but they are never edited.
- Dates as a distinct type (they are read as Excel serial numbers), volatile functions, array or spill formulas, structured table references, R1C1, external links. These are treated as opaque.
- Relocating merged cells, tables, conditional formatting and data validation. Structural edits are refused on sheets that contain them.
- Moving or copying ranges, sorting and filtering.
- **Stretch only:** formula bytecode VM with an AST optimiser (the Revision 1 Phase 3 plan); an MCP server wrapper for the CLI; an Office Scripts code-generation target.

---

## 6. Proposed Solution

```
                 ┌──────────────────────── agent (any LLM with a shell) ─────────────────────────┐
                 │   tabula inspect book.xlsx          tabula check book.xlsx edits.tel            │
                 └───────────────┬────────────────────────────────┬──────────────────────────────┘
                                 ▼                                ▼
   TEL compiler:  lexer ─► parser ─► resolver ─► lowering ─► IR opt ─► simulate + analyse ─► plan
   (outer)          │ formula mode                 (fill)    (dead writes)  (cycles, impact,   │ text / JSON
                    ▼                                                          expects)        ▼
   Formula compiler: lexer ─► Pratt parser ─► binder/checker ─► evaluator · emitter   apply: back end
   (inner)                                                                              (replay + sync)
                    ▼                                                                          ▼
   Workbook engine:  cells · defined names · dependency graph · cycles · incremental recompute   .xlsx
                    ▲                                                                          ▲
   I/O boundary:     openpyxl load (formulas + Excel-cached values)           atomic save + verify + fidelity guard
```

**Originality claims**

1. **Static semantics for agent edits.** The workbook is the symbol table. Every reference in an agent's script is resolved and checked before execution, and sheet-name typos get edit-distance suggestions.
2. **Pre-execution impact analysis.** The exact write-set and the downstream cells it affects (blast radius) come from the dependency graph, together with predicted before/after values.
3. **Reference relocation for structural edits** across formulas, sheets and defined names, matching Excel's rules (including `#REF!` on deletion). openpyxl does not do this.
4. **Optimisation and lint over an edit IR.** Dead-write elimination, a formula-overwritten-by-constant ("hard-coding") warning, and cycle rejection.
5. **Honest partial evaluation.** Formulas Tabula cannot simulate are opaque, and their uncertainty propagates as *unverified* rather than being guessed.
6. **A controlled A/B evaluation** of agent editing with Tabula vs. the default method.

---

## 7. Functional Requirements

### Formula language and engine

| ID | Requirement |
|----|-------------|
| FR-1 | The engine shall model a workbook of named sheets (case-insensitive names) with cells addressable up to `XFD1048576`, each holding a literal or a formula, plus defined names at workbook and sheet scope. |
| FR-2 | The formula lexer shall tokenize the Excel subset of §5, including sheet-qualified (quoted and unquoted), absolute, whole-column and whole-row references, postfix `%`, and error literals, reporting positioned lexical errors. |
| FR-3 | The formula parser shall implement Excel precedence: comparison < `&` < `+ -` < `* /` < `^` (**left**-associative) < `%` < unary `-`. It shall build an AST and report positioned syntax errors. |
| FR-4 | Formula semantic analysis shall resolve names (TEL `let` scope → sheet-scoped defined name → workbook defined name), check arity and argument kinds for supported functions, classify other calls as *known-unsimulated* (in the Excel function catalogue) or *unknown* (error), and reject out-of-bounds references. |
| FR-5 | The formula emitter shall regenerate Excel formula text from an AST, quoting sheet names when needed, such that `parse(emit(ast)) == ast`. |
| FR-6 | Evaluation shall follow Excel semantics for the supported subset: numeric-text coercion, case-insensitive text comparison, cross-type ordering, aggregates that skip text in ranges, `ROUND` half away from zero, and the error values `#DIV/0! #VALUE! #REF! #NAME? #N/A #NUM!`. |
| FR-7 | The engine shall maintain a dependency graph with cell edges and symbolic range edges across sheets, and answer precedent/dependent queries. |
| FR-8 | The engine shall detect circular references and report the cycle path. |
| FR-9 | The engine shall recompute only the transitive dependents of changed cells, in topological order, and report recompute statistics. |
| FR-10 | Formula cells using functions the engine does not simulate, or that it cannot parse, shall keep their Excel-cached value. Their values shall be marked **unverified** whenever an input may have changed, and the mark shall propagate to dependents. |

### Workbook I/O

| ID | Requirement |
|----|-------------|
| FR-11 | The system shall load `.xlsx` workbooks with formulas and Excel-cached values, translating dates to serial numbers and stripping storage prefixes such as `_xlfn.`. |
| FR-12 | `inspect` shall report sheets, used ranges, header rows, defined names, function usage, and **engine agreement**: how many formulas Tabula evaluates to the same value as Excel's cache. |
| FR-13 | Saving shall be atomic (write to a temporary file, then rename), verified (reload and compare every written cell, and after structural edits every cell), and guarded. It is refused when the saved file would lose workbook parts present in the input (fidelity guard), when source or output is open in Excel, when `--if-unchanged` does not match the file's SHA-256, or when source or output changes during apply. Cooperating writers are serialised by advisory locks. |

### TEL compiler

| ID | Requirement |
|----|-------------|
| FR-14 | The TEL lexer shall tokenize scripts line by line, switching to formula mode for the formula part of `set … =` and `expect`, so that formula diagnostics carry positions in the script. |
| FR-15 | The TEL parser shall accept the statements of §5 and the grammar in design.md §4, reporting positioned syntax errors and continuing at the next line (statement-level error recovery). |
| FR-16 | The resolver shall maintain a scope chain (`in` blocks, `let`), detect redeclaration and use-before-declaration, resolve sheets (with edit-distance suggestions) and names, require sheet qualification outside `in` blocks in multi-sheet workbooks, and check list shapes against target shapes and target sizes against limits. |
| FR-17 | Lowering shall produce an edit IR of per-cell operations. A formula written to a range shall be filled with relative references shifted per cell (Excel fill semantics), and a fill that pushes a reference off the grid is an error. |
| FR-18 | The IR optimiser shall remove dead writes (a cell written again before any read barrier) and report each as a warning. |
| FR-19 | Analysis shall compute the write-set, enforce the `--allow` allowlist, warn when a constant overwrites a formula, and reject edits that create circular references (with path). |
| FR-20 | Simulation shall apply the IR to the engine and recompute. It shall report predicted values for written and affected cells, cells that newly evaluate to errors, unverified cells, and the outcome of every `expect`; a false or error-valued `expect` is an error. |
| FR-21 | Structural statements shall relocate references in all formulas, defined names and pending `expect`s exactly as Excel does: shift, expand, shrink, or `#REF!` on deletion. They (and sheet renames) shall be refused when any sheet has address-bearing features Tabula does not relocate, or any unparsed or unsupported formula, and when an insert would push content off the grid. |
| FR-22 | `check` shall never modify the workbook. `apply` shall write only when compilation produced no errors. |
| FR-23 | Plans shall be rendered as text or JSON (schema version `1`), with deterministic ordering and bounded list sizes (counts always given). |

### Interface and demo

| ID | Requirement |
|----|-------------|
| FR-24 | The CLI shall provide `inspect`, `deps`, `explain`, `check`, `apply` and `repl`, with exit codes 0 (success), 1 (diagnostics/errors), and 3 (usage or I/O error). |
| FR-25 | The demo shall run each task under both methods (Tabula, default openpyxl) with an isolated agent session per run, record transcripts, cost and duration, score each output with value-based checkers that are independent of Tabula's engine, and produce a comparison report. |
| FR-26 | No input shall crash the system. Every failure surfaces as a diagnostic. |

## 8. Non-Functional Requirements

| ID | Requirement |
|----|-------------|
| NFR-1 | **Portability:** Python ≥ 3.11; runtime dependency **openpyxl** (pinned), imported only by the I/O boundary module; everything else uses the standard library. pytest is dev-only. The demo checker uses an independent evaluator (pycel, which needs openpyxl 3.0) in a separate demo-only virtualenv. |
| NFR-2 | **Performance:** on a 10,000-formula workbook, simulating a bounded edit shall take < 100 ms excluding file I/O, and a full `check` including load shall take < 5 s. |
| NFR-3 | **Modularity:** formula compiler, engine, TEL compiler and I/O boundary communicate only through defined types (AST, Workbook, IR, Plan). |
| NFR-4 | **Testability:** every stage is unit-testable without files. Workbook-level behaviour is covered by scenario and golden tests. |
| NFR-5 | **Code quality:** meaningful names, docstrings on public interfaces, comments on non-obvious algorithms (manual §17, §22). |
| NFR-6 | **Data safety:** `check` never writes. `apply` never leaves a partially written file and never writes after an error. |
| NFR-7 | **Agent-facing output:** every diagnostic has a stable code, a position and (where possible) a hint. JSON output is versioned and its size is bounded. |
| NFR-8 | **Determinism:** identical inputs produce byte-identical plans. |

## 9. Input / Output Requirements

- **Input:** `.xlsx` workbooks; TEL scripts (`.tel`, UTF-8); CLI flags; REPL commands.
- **Output:** plans and diagnostics (text or JSON) on stdout; modified `.xlsx` files (with `apply`); demo run folders (transcripts, output workbooks, `REPORT.md`).

---

## 10. Compiler Design Concepts Involved

| Course Concept | Where it appears in Tabula |
|---|---|
| Lexical analysis | Formula scanner (references, sheet qualifiers, error literals); TEL scanner with **lexer modes** (formula island inside TEL) |
| Grammars, EBNF | Formula grammar (Excel precedence, left-assoc `^`, postfix `%`); TEL statement grammar |
| Parsing | Pratt parser for formulas; recursive-descent statement parser for TEL with statement-level error recovery |
| Symbol tables, scope | Workbook symbol table (sheets, defined names with sheet → workbook scope); TEL scope chain (`in` blocks, `let`), shadowing, redeclaration and use-before-declaration checks |
| Semantic analysis, type checking | Name resolution, arity/kind checks, list-vs-range **shape checking**, target-size limits, sheet-qualification rule, cycle detection as a whole-program check |
| Intermediate representation | Edit IR: a linear list of per-cell quadruple-like operations (`SET_FORMULA sheet cell ast`, `INSERT_ROWS sheet at n`, …) |
| Optimisation / dataflow | Dead-write (dead-store) elimination over the IR; dependency closure (blast radius); unverified-value (taint) propagation; incremental recomputation |
| Code generation | Formula emitter (AST → Excel text); back end lowering IR to openpyxl operations (replay structure, sync contents) |
| Relocation | Fill (relative-reference shifting) and insert/delete relocation: the linker's relocation problem applied to cell addresses |
| Interpretation and runtime | Formula evaluator with Excel error-value semantics; simulator executing the IR against the engine |
| Error detection and recovery | Positioned, coded diagnostics in both languages; statement-level recovery in TEL; per-cell containment in the engine; transactional, all-or-nothing writes |
| Scheduling | Kahn's algorithm over the dirty subgraph |

## 11. Tools and Technologies

| Category | Choice | Justification |
|---|---|---|
| Language | Python 3.11+ | Recommended by the manual (§20); suits tree- and graph-heavy code |
| Parsers | Hand-written (no Lex/Yacc/ANTLR) | Every algorithm is explainable in the viva; full control of positioned diagnostics |
| File format | openpyxl 3.1.x (I/O only) | Reads and writes OOXML; acknowledged per manual §18. Does not evaluate formulas or relocate references; Tabula implements both |
| Testing | pytest | Unit, scenario, golden, differential, benchmark |
| Demo agent | Claude Code (headless `claude -p`, isolated `--safe-mode`) | A real agent with a shell, driven identically in both arms |
| Demo checker | pycel (demo-only virtualenv) | Independent formula evaluator, so the demo does not grade itself with Tabula's own engine |
| Version control | Git | Manual §23 |

## 12. Deliverables Mapped to Course Phases

| Phase | Course deliverable | Artifact |
|---|---|---|
| Phase 1 (done, Review 1) | Proposal, design, language spec, prototype | Revision 1 docs; formula lexer, parser, evaluator, REPL; 54 tests |
| Phase 2 (Review 2) | Core modules, working code, demo, tests, module explanation, progress report | Excel-compatible formula front end + emitter; workbook engine (graph, cycles, recompute); `.xlsx` I/O; TEL front end, resolver, IR, simulation; `check`/`apply` for cell edits; first A/B demo |
| Phase 3 (Review 3) | Complete system, full tests, results, report, presentation, user instructions | Structural edits + relocation; dead-write elimination; guards (allowlist, concurrency, fidelity); JSON schema; benchmarks; differential corpus; full A/B evaluation with results; final report |

## 13. Expected Outcomes

1. A CLI that an agent can use to inspect a workbook, check an edit script, and apply it safely, with diagnostics specific enough for the agent to correct itself.
2. Demonstrated detection, before any write, of the agent failure classes: hallucinated sheets and names, wrong shapes, hard-coding over formulas, circular references, broken references after structural edits, and violated expectations.
3. Measured results:
   - A/B task success and silent-failure rates, Tabula vs. default method;
   - engine agreement with Excel-cached values;
   - incremental vs. full recompute on large workbooks.
4. Complete documentation per the manual (§24).

## 14. Constraints and Assumptions

- Individual work. AI tools are used per manual §19, and everything must be explainable and modifiable by the student.
- No `eval`. All formula evaluation goes through Tabula's evaluator.
- Grid bounds are Excel's: `A`–`XFD` × `1`–`1,048,576`.
- Numbers are Python floats, displayed in general format. Dates are Excel serial numbers.
- **Faculty approval of this scope extension is required (manual §4)** before Review 2. §15 is written to support that conversation.

## 15. Changes from Revision 1 (Review 1)

| Revision 1 | Revision 2 | Why |
|---|---|---|
| "Tabula is its own language, not an Excel clone" | Excel-compatible formula subset | Tabula now writes formulas into real workbooks; any dialect difference would be a bug when Excel opens the file |
| `^` right-associative (`=2^3^2` → 512) | **Left-associative** (`=2^3^2` → 64) | That is what Excel computes |
| `%` infix modulo; `AVG` | Postfix percent (`50%` → 0.5); `MOD(a,b)`; `AVERAGE` | Excel syntax |
| Text never numeric (`="2"+1` → `#VALUE!`) | Numeric text coerces (`="2"+1` → 3) | Excel semantics |
| Grid A–ZZ × 10,000 | A–XFD × 1,048,576 | Excel bounds |
| TSV load/save; REPL as main interface | `.xlsx` I/O; CLI with JSON for agents; REPL kept | Agents edit real files |
| Cross-sheet references, insert/delete rows out of scope | In scope | Essential for real workbooks, and the core of the relocation work |
| Ranges expanded to cells (cap 50k) | Symbolic range edges | Whole-column ranges (`A:A`) cover 1,048,576 cells |
| Recursive prototype evaluator | Topological evaluation (no recursion) | Real workbooks have long dependency chains |
| Phase 3: bytecode VM + AST optimiser + HTML report | Stretch. Replaced by edit IR + dead-write elimination + emitter/back end | The edit IR, its optimiser and the back end cover the IR, optimisation and code-generation concepts with practical payoff |
| Standard library only | openpyxl at the I/O boundary only | Reading and writing OOXML is not a compiler concern; manual §18 allows acknowledged libraries |
