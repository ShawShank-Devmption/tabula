# Design Document — Software Requirements Specification & Architecture

**Project:** Tabula — A Statically Checked Spreadsheet Edit Language for AI Agents, built on a Compiled Formula Engine

**Revision 2 (2026-10-08), safety hardening of the same date (§6.3, §7, §8.6).** `requirements.md` says *what* and *why*; this document says *how*. Revision 1 (Review 1) specified the formula compiler and sheet engine. Revision 2 makes them Excel-compatible and multi-sheet, and adds the TEL compiler, the `.xlsx` boundary and the A/B demo.

---

## 1. System Architecture

### 1.1 Layers

```
 ┌──────────────────────────── CLI (tabula/cli.py) ─────────────────────────────┐
 │ inspect · deps · explain · check · apply · repl          text | --json       │
 └───────────────┬──────────────────────────────────────────────┬───────────────┘
                 ▼                                              ▼
 ┌──────── TEL compiler (tabula/tel/) ────────────────────────────────────────────┐
 │ lexer ─► parser ─► resolver ─► lowering ─► IR optimiser ─► simulate/analyse ─► │
 │ (formula   (TEL AST)  (scopes,   (per-cell   (dead writes)   (cycles, impact,  │
 │  mode)                 checks)    edit IR)                     expects)   plan  │
 └───────┬─────────────────┬───────────────────────────────┬──────────────┬──────┘
         │ formula text    │ formula AST                    │ IR ops       │ apply
         ▼                 ▼                                ▼              ▼
 ┌── Formula compiler ─────────────────┐   ┌── Workbook engine ───────┐  ┌── I/O boundary ──┐
 │ lexer → Pratt parser → AST          │   │ Workbook/Sheet/Cell       │  │ tabula/xlsx.py   │
 │ functions (registry + catalogue)    │──►│ DependencyGraph           │◄─│ load (formulas + │
 │ evaluator (Excel semantics)         │   │ Engine: recompute (Kahn,  │  │  cached values)  │
 │ emitter (AST → Excel text)          │   │  Tarjan), taint           │  │ write_back       │
 │ relocate (fill, insert/delete, ren.)│   └───────────────────────────┘  │ (replay + sync,  │
 └─────────────────────────────────────┘                                   │  fidelity, atomic│
                                                                           │  verify)         │
                                                                           └──────────────────┘
```

**Dependency rule.**
- The formula compiler depends only on `refs` and `values`.
- The engine depends on the formula compiler.
- The TEL compiler depends on the engine and the formula compiler.
- Only `tabula/xlsx.py` imports openpyxl.
- `cli.py` depends on everything, and nothing depends on it.

### 1.2 Module Decomposition

```
tabula/
├── __main__.py        python3 -m tabula → cli.main
├── refs.py            CellRef, RangeRef (cells | cols | rows), A1 parse/format, Excel bounds,
│                      sheet-name quoting
├── values.py          value model, ErrorValue, Excel coercions and cross-type comparison
├── lexer.py           formula scanner (sheet qualifiers, error literals, $A / $1 parts)
├── parser.py          formula Pratt parser, AST, AST dump
├── emitter.py         AST → Excel formula text
├── functions.py       supported FunctionSpecs (simulated) + EXCEL_FUNCTIONS catalogue
├── evaluator.py       evaluate(ast, host, ctx) — one formula, Excel semantics
├── relocate.py        fill shifting, insert/delete relocation, sheet rename over ASTs
├── workbook.py        Workbook, Sheet, Cell, DefinedName; formula precedents
├── graph.py           DependencyGraph: cell + range edges, closure, Kahn, Tarjan, cycle path
├── engine.py          Engine (recompute, taint, stats, expression evaluation) + REPL Sheet facade
├── diagnostics.py     Diagnostic, caret/text and JSON rendering, edit-distance suggestions
├── xlsx.py            openpyxl boundary: snapshot load, write_back, fidelity guard, writer locks,
│                      commit-time revalidation, sha256
├── tel/
│   ├── lexer.py       line scanner with parser-controlled modes
│   ├── parser.py      TEL AST + recursive-descent statement parser with line recovery
│   ├── resolver.py    scope chain, symbol tables, semantic checks, lowering to IR
│   ├── ir.py          IR op types, dead-write elimination
│   ├── simulate.py    executes IR on the engine; analyses; expects
│   ├── plan.py        Plan type; text and JSON (schema v1) rendering
│   └── compiler.py    check() / apply() pipelines
└── cli.py             subcommands
demo/                  A/B agent demo (§10)
docs/TEL.md            TEL language reference (also handed to the demo agent)
```

---

## 2. External Interfaces

### 2.1 CLI

```
python3 -m tabula inspect BOOK.xlsx [RANGE] [--json]     sheets, headers, names, engine agreement;
                                                         or a range: raw content + value per cell
python3 -m tabula deps    BOOK.xlsx CELL [--json]        precedents, dependents (direct + transitive)
python3 -m tabula explain BOOK.xlsx CELL                 formula pipeline: tokens, AST, refs, value
python3 -m tabula check   BOOK.xlsx SCRIPT.tel [--json] [--allow SPEC]...
                                                         compile + simulate; NEVER writes
python3 -m tabula apply   BOOK.xlsx SCRIPT.tel (-o OUT.xlsx | --in-place) [--json]
                          [--allow SPEC]... [--if-unchanged SHA256] [--allow-lossy]
python3 -m tabula repl    [SCRIPT.txt]                   Phase 1 REPL (single in-memory sheet)
```

- `CELL` and `RANGE` are written as `Sheet!A1` or `'Q3 Sales'!A1:F10`. A bare `A1` is accepted only for single-sheet workbooks.
- `--allow SPEC` adds a write allowlist entry: `Sheet` (whole sheet) or `Sheet!A1:F40`. Once any entry is given, every write must fall inside an entry. Structural edits need the whole sheet. `add sheet` and `rename sheet` need the entry `*`.
- **Exit codes:** 0 success (warnings allowed) · 1 the script has errors or `apply` was refused · 3 usage or I/O error.

### 2.2 Diagnostics

Every diagnostic has a severity, a stable code, a 1-based position in the script, a message and an optional hint:

```
error[E-SHEET] edits.tel:3:5: unknown sheet 'Summry'
   3 | set Summry!B8 =Sales!G22
     |     ^^^^^^
  hint: did you mean 'Summary'?
```

| Code | Meaning |
|---|---|
| `E-LEX`, `E-SYNTAX` | lexical / syntax error (TEL or embedded formula) |
| `E-SHEET` | unknown sheet (with suggestion) |
| `E-NOSHEET` | unqualified reference outside an `in` block in a multi-sheet workbook |
| `E-NAME` | unknown `let` name or defined name (with suggestion) |
| `E-FUNC` | function unknown to Excel (with suggestion) |
| `E-ARITY` | wrong number of arguments to a supported function |
| `E-TYPE` | range used where a single value is needed; `let` value used as a target |
| `E-SHAPE` | list shape ≠ target shape; ragged list |
| `E-DECL` | `let` redeclared in the same scope |
| `E-SIZE` | target larger than 100,000 cells |
| `E-REF` | reference outside the grid (including after fill shifting) |
| `E-CYCLE` | the edit creates a circular reference (path given) |
| `E-EXPECT` | an `expect` is FALSE, an error, or not boolean |
| `E-PERM` | write outside the `--allow` list |
| `E-STRUCT` | insert/delete/rename refused, because some stored address might not be relocated. Applies **workbook-wide**: any sheet with charts, pivot tables, chart sheets, external links, merged cells, tables, conditional formatting, data validation, hyperlinks, a print area/titles or an autofilter; any unparsed formula or name; or any formula/name using functions Tabula does not simulate (e.g. `INDIRECT`, whose text addresses cannot be relocated). Also refused when an insert would push stored cells or formatting off the grid |
| `E-FIDELITY` | saving would lose workbook parts (apply refused) |
| `E-LOCKED` | workbook (source or output) is open in Excel, before or during apply; or the platform lacks POSIX advisory locks |
| `E-CONFLICT` | `--if-unchanged` hash mismatch, or the source/output file changed while apply was running |
| `E-VERIFY` | the written file did not read back as intended (apply refused, input untouched) |
| `E-WRITE` | the output could not be written, or the structural edits could not be replayed (input untouched) |
| `W-OVERWRITE-FORMULA` | a constant replaces an existing formula (hard-coding) |
| `W-DEAD-WRITE` | a write is overwritten later and has no effect |
| `W-NEW-ERROR` | a cell that had a value now evaluates to an error |
| `W-UNSIMULATED` | the formula uses an Excel function Tabula does not simulate; value unverified |
| `W-UNPARSED` | the workbook has formulas Tabula cannot parse; impact may be incomplete |
| `W-SCOPE` | formula written to another sheet uses unqualified references (they refer to the target's sheet) |
| `W-SHADOW` | `let` shadows an outer name |
| `W-STALE-LET` | a `let`-bound address is used after a structural edit on its sheet |
| `W-EXPECT-UNVERIFIED` | an `expect` depends on unverified values or unsupported functions; it is reported as unverified, never as passed |
| `W-LOSSY` | (check) saving this workbook would lose parts; apply needs `--allow-lossy` |

### 2.3 Plan Output

The text form is for humans and agents. The JSON form (`--json`) uses `"schema": 1`:

```json
{ "schema": 1, "ok": true, "applied": false, "workbook": "book.xlsx", "sha256": "…",
  "diagnostics": [ {"severity":"warning","code":"W-DEAD-WRITE","line":7,"col":1,
                    "message":"…","hint":"…"} ],
  "writes":   {"count": 42, "items": [ {"cell":"Sales!G2","before":{"raw":"","value":null},
                                         "after":{"raw":"=E2*TaxRate","value":1260.0},
                                         "unverified": false} ]},
  "affected": {"count": 3,  "items": [ … same shape … ]},
  "new_errors": {"count": 0, "items": []},
  "unverified": {"count": 0, "items": []},
  "expects":  [ {"line": 9, "source": "Summary!B8>0", "status": "pass"} ],
  "stats": {"statements": 6, "ops": 44, "dead_writes": 0, "cells_written": 42,
            "formulas_recomputed": 45, "total_formulas": 61, "structural": false,
            "recompute_ms": 2.9} }
```

Lists are capped (`writes`/`affected`: 50 items in JSON, 12 in text) and `count` is always exact. Ordering is deterministic: sheet order, then row, then column.

---

## 3. Formula Language — Lexical Level (Excel subset)

| Class | Pattern / members |
|---|---|
| Cell reference | `\$?[A-Z]{1,3}\$?[0-9]+`, column ≤ `XFD`, 1 ≤ row ≤ 1,048,576 |
| Column part | `\$?[A-Z]{1,3}` (only inside `A:C`) |
| Row part | `\$?[0-9]+` (only inside `1:3`) |
| Sheet qualifier | `Name!` or `'Any name'!` (`''` escapes a quote) |
| Number | `[0-9]+(\.[0-9]*)?([eE][+-]?[0-9]+)?` or `\.[0-9]+…` |
| String | `"…"` with `""` as an escaped quote |
| Boolean | `TRUE` `FALSE` |
| Error literal | `#DIV/0!` `#VALUE!` `#REF!` `#NAME?` `#N/A` `#NUM!` `#NULL!` |
| Identifier | `[A-Za-z_\\][A-Za-z0-9_.]*` — function or defined name. Storage prefixes `_xlfn.` and `_xlws.` are stripped |
| Operators | `+ - * / ^ % &  = <> < <= > >=` |
| Delimiters | `( ) , :` |

**Classification rules.**
- A word followed by `!` is a sheet qualifier, even if it looks like a cell reference (`Q3!A1`).
- A word that matches the cell-reference pattern is a reference.
- `TRUE`/`FALSE` are booleans.
- A word containing `$` that is not a cell reference is a column part.
- Anything else is an identifier.
- Pure-letter words such as `SUM`, `TAX` or `A` are ambiguous between function, name and column. The parser decides by context (`(` means a call, `:` plus a column means a column range, otherwise a name).
- Array constants `{…}`, the `@` intersection operator and `#` spill references are lexical errors. When they appear in loaded formulas, those cells become *unparsed* (opaque).

## 4. Formula Language — Grammar (EBNF)

```ebnf
formula     = [ "=" ] expression EOF ;
expression  = comparison ;
comparison  = concat   { ( "=" | "<>" | "<" | "<=" | ">" | ">=" ) concat } ;
concat      = additive { "&" additive } ;
additive    = mult     { ( "+" | "-" ) mult } ;
mult        = power    { ( "*" | "/" ) power } ;
power       = postfix  { "^" postfix } ;              (* LEFT-associative, as in Excel *)
postfix     = prefix   { "%" } ;
prefix      = ( "-" | "+" ) prefix | primary ;
primary     = NUMBER | STRING | BOOLEAN | ERROR
            | [ SHEET ] reference
            | IDENT "(" [ expression { "," expression } ] ")"
            | "(" expression ")" ;
reference   = CELL [ ":" CELL ] | COL ":" COL | ROW ":" ROW | IDENT     (* IDENT = defined name *)
            | ERROR ;                                                    (* Sheet!#REF! *)
```

Precedence, loosest to tightest: comparison < `&` < `+ -` < `* /` < `^` < postfix `%` < prefix `-`. So `=-2^2` is 4, `=2^3^2` is 64, and `=50%` is 0.5. These are Excel's rules, and they **change Revision 1**, where `^` was right-associative.

**AST nodes.** `Number Text Bool ErrorLit Ref(ref, sheet) RangeNode(rng, sheet) Name Unary Percent Binary Call`. `sheet` is the qualifier as written (`None` means the host cell's sheet). `RangeNode.rng.kind` is one of `cells | cols | rows`.

### 4.1 Static Semantics

- **Name resolution:** TEL `let` scope chain (innermost first) → sheet-scoped defined name of the host sheet → workbook defined name → `E-NAME`. Inside formulas, a `let` name is **inlined**: it is replaced by its reference or literal before emission, because Excel cannot see TEL names. A defined name stays a `Name` node, since Excel resolves it.
- **Functions:**
  - A name in the supported registry (§5.2) is checked for arity and argument kinds.
  - A name in the Excel catalogue but not the registry is *known-unsimulated* (`W-UNSIMULATED`).
  - Any other name is `E-FUNC`, with an edit-distance suggestion.
- **Ranges are not values.** A range is legal only as an argument to a range-accepting parameter. Elsewhere it is `E-TYPE`, because Excel's implicit intersection and spill behaviours are not supported.
- **Bounds:** every reference must lie in `A1:XFD1048576` (`E-REF`).

### 4.2 Value Semantics (Excel-compatible)

| Rule | Behaviour |
|---|---|
| Numeric context | Blank → 0; `TRUE`/`FALSE` → 1/0; numeric text → number (`"2"+1` → 3); other text → `#VALUE!` |
| Text context (`&`) | Blank → `""`; numbers in general format (`3`, `0.333333333333333`); `TRUE`/`FALSE` |
| Comparison | Numbers < text < booleans across types. Text compares case-insensitively. Numbers compare at 15 significant digits (`0.1+0.2=0.3` is TRUE). Blank takes the other side's type (0, `""`, FALSE) |
| Aggregates over references | `SUM AVERAGE MIN MAX` use only the numeric cells of any *reference*, whether a range or a single cell (text, booleans and blanks are skipped). Errors propagate. Typed scalar arguments are coerced. `MIN`/`MAX` of nothing → 0; `AVERAGE` of nothing → `#DIV/0!` |
| `COUNT` / `COUNTA` | Numbers only / all non-blank cells; errors never propagate |
| `COUNTIF` / `SUMIF` | Criterion is a number, text (case-insensitive, `*` and `?` wildcards) or an operator prefix (`">=10"`, `"<>North"`). Blank cells count where Excel counts them (`""` and `"<>x"` match blanks) |
| `ROUND` | Half away from zero (`ROUND(2.5,0)` → 3), unlike Python's `round` |
| `IF` / `IFERROR` / `AND` / `OR` | `IF` is lazy (only the taken branch is evaluated); a 2-argument `IF` yields FALSE. `AND`/`OR` over ranges skip text and blanks |
| Errors | `#DIV/0!` (÷0, `MOD` by 0), `#VALUE!`, `#REF!`, `#NAME?`, `#N/A`, `#NUM!` (`SQRT` of a negative, `0^0`, overflow: Excel has no infinities). The first error in evaluation order wins. `#CYCLE!` is internal to the simulator and never written |

These rules are tested against Excel-cached values in real workbooks (§11).

## 5. Formula Compiler Modules

### 5.1 Lexer and Pratt parser
The single-pass scanner uses maximal munch. Binding powers are comparison 10, `&` 20, `+ -` 30, `* /` 40 and `^` 50, all left-associative. Prefix `-`/`+` and postfix `%` are parsed below the Pratt loop.

### 5.2 Function registry and catalogue
`FunctionSpec(name, min_args, max_args, range_params)` covers the 27 simulated functions in requirements §5. `range_params` is `"all"` or a set of argument indexes. `EXCEL_FUNCTIONS` is a catalogue of about 400 Excel function names, used only to separate real-but-unsimulated functions from hallucinated ones.

### 5.3 Emitter
`emit(ast, host_sheet)` regenerates Excel text:
- minimal parentheses from the binding powers (a child is wrapped when its power is lower than its parent's, or equal on the right-hand side);
- sheet names quoted when they are not plain identifiers or when they look like references (`'Q3 Sales'!A1`, `'Q3'!A1`);
- numbers in general format.

The test property is `parse(emit(parse(s))) == parse(s)`.

When writing to a file, a function added to Excel after 2007 is emitted with Excel's storage prefix (`_xlfn.CONCAT`, `_xlfn._xlws.SORT`); without it Excel shows `#NAME?`. The lexer strips these prefixes on load.

### 5.4 Evaluator
`evaluate(ast, host_sheet, ctx)` walks one AST. It never recurses into other cells. `ctx` supplies the current cell values, iterates ranges (stored cells only, so `A:A` costs the sheet's size, not 1,048,576), and resolves defined names. This removes Revision 1's recursion-depth limit: deep chains are handled by the engine's topological order (§6.3).

### 5.5 Relocation (`relocate.py`)
There are three pure transforms over an AST. Each returns `(new_ast, changed)`.

| Transform | Rule |
|---|---|
| `shift_for_fill(ast, dc, dr)` | Relative parts move by the offset and `$` parts stay. A result off the grid is an error. This is Excel's fill/copy semantics |
| `relocate(ast, host, sheet, axis, at, n)` | **Insert** (`n > 0`): positions ≥ `at` move by `n`, whether absolute or not. A range spanning `at` (start < at ≤ end) **expands**. A range ending at `at - 1` does not, which is Excel behaviour. **Delete** (`n < 0`, span `at … at+|n|-1`): positions after the span move back. A range partially covered **shrinks**. A cell or range wholly inside the span becomes `#REF!`. Only references whose sheet resolves to the edited sheet are touched |
| `rename_sheet(ast, old, new)` | Explicit qualifiers equal to `old` (case-insensitive) become `new` |

The same functions relocate defined names and pending `expect`s. This is the linker's relocation problem: every stored address is rewritten when the address space changes.

## 6. Workbook Engine

### 6.1 Model
```
Cell     uid, raw, kind (literal | formula | unparsed), ast, value, cached,
         simulable, unverified, touched, text_changed
Sheet    sid (stable), name, cells {(col,row): Cell}, meta {merged, tables, cf, dv}
Workbook sheets (ordered), names {(scope_sid | None, NAME): DefinedName(ast)}
```
- Cell keys are `(sid, col, row)`. Sheet ids are stable, so a rename touches only the name table and the qualifiers in ASTs.
- `uid` follows a cell through moves, which lets before/after be compared across structural edits.

### 6.2 Dependency graph
- **Cell edges:** `precedent key → {dependent formula keys}`.
- **Range edges:** stored *symbolically*, per sheet, as `(bounds, dependent)`. They are not expanded, so whole-column references stay cheap.
- `SUMIF(criteria, crit, sum)` reads the **effective** sum range: `sum`'s top-left cell sized like `criteria`, as Excel does. That range, not just the cells written in the formula, is a precedent edge, so it takes part in invalidation and in cycle detection.
- `dependents(k)` returns the direct cell edges plus every range edge whose bounds contain `k`. Range edges are bucketed by column, so a lookup scans only ranges over `k`'s column. Ranges wider than 64 columns sit in a small per-sheet list.
- Formula-to-formula precedents within a set use `formulas_in(sid, bounds)`. That iterates the smaller of the range and the sheet's formula index.

| Operation | Algorithm |
|---|---|
| `closure(seeds)` | BFS over `dependents` |
| `topo_order(D)` | Kahn's algorithm restricted to `D`, with in-degrees counted inside `D` |
| cycles | Nodes Kahn cannot schedule go to an iterative **Tarjan SCC**. Members of a non-trivial SCC (or a self-loop) are marked `#CYCLE!`. The remainder is scheduled again with Kahn |
| `cycle_path(k)` | Iterative DFS returning `k → … → k` for diagnostics |

### 6.3 Recompute
```
recompute(seeds):                       seeds = None → all formulas
  D     = closure(seeds) ∩ formulas
  order = topo_order(D)                 cycle members → #CYCLE!
  for f in order:
     if f.simulable: f.value = evaluate(f.ast, …)
     else:           f.unverified = True           (value = last cached, shown as "?")
     if any precedent of f is unverified: f.unverified = True       (taint propagation)
  stats: |D|, total formulas, elapsed ms
```
- At load, simulable formulas are computed by Tabula and non-simulable ones keep Excel's cached value; both are considered verified.
- **Conservative invalidation.** Once anything is edited, *every* unparsed formula and every formula using an unsupported function or name is marked unverified, together with its dependents. Functions like `INDIRECT` or `OFFSET` can read cells that no AST edge shows, so missing edges are not proof that a cached value is still right. An `expect` over such values is reported unverified.
- **Engine agreement** is the number of simulable formulas whose Tabula value equals the cached value (relative tolerance 1e-9).

## 7. I/O Boundary (`xlsx.py`)

**Load.**
- The file is read **once into memory**. Both openpyxl loads (formulas, then `data_only=True` for cached values), the SHA-256 and the fidelity part counts all come from that one byte snapshot, so a save by another program mid-load cannot mix two versions.
- Formulas are parsed. If parsing fails, the cell is `unparsed` (value = cached, no edges). Array and data-table formulas are also `unparsed`.
- Dates become Excel serial numbers in the workbook's own date system (1900 or 1904 epoch). Defined names are parsed at workbook and sheet scope.
- Per-sheet `meta` counts the address-bearing features Tabula does not relocate (see `E-STRUCT`). Stored cell positions (including style-only cells), row heights and column widths are recorded too, so grid overflow can be checked.
- openpyxl warnings during load are captured; they signal lossy features.

**Write back** (`apply` only, after a clean compile):
1. **Replay structure.** For each structural or sheet op in program order: `insert_rows / delete_rows / insert_cols / delete_cols / create_sheet / title = new`. openpyxl moves cells and styles but rewrites no formulas and does not move row heights, hidden flags or column widths; Tabula shifts those dimensions itself. A rename that changes only letter case goes through a temporary name, because openpyxl would otherwise de-duplicate it.
2. **Sync contents.** Every model cell flagged `touched` or `text_changed` is written at its final address: a literal; `None` for cleared cells (an empty-string literal clears); or `"=" + emit(ast)` with storage prefixes. Literal text that starts with `=` or looks like an error code (`"#N/A"`) gets `data_type 's'`, or openpyxl would store it as a formula or an error. Relocated defined names are rewritten too.
3. **Save to a temporary file** in the same directory.
4. **Fidelity guard.** Count zip parts per category (charts, drawings, media, pivot tables/caches, tables, comments, threaded comments, VBA, slicers, timelines, external links, controls, embeddings, custom XML) in the input and in the temporary file. Any decrease, or a lossy load warning, gives `E-FIDELITY` unless `--allow-lossy` is set.
5. **Verify.** Reload the temporary file and compare every synced cell with the model. After a structural edit, check **every** model cell (including untouched values that merely moved) and reject any unexpected non-empty cell.
6. **Revalidate, then commit.** Re-check Excel owner files for source and output, re-hash the source against the loaded snapshot, and check the output still matches its state at the start. Then `os.replace(temp, output)`. On any failure the temporary file is deleted and the input is untouched.

**Guards:**
- `apply` holds advisory `fcntl` locks on sidecar files (`.<name>.xlsx.tabula.lock`, kept on purpose) for source and output, acquired in sorted order. They serialise cooperating Tabula writers.
- Excel lock file `~$name` (before apply and again just before commit) → `E-LOCKED`.
- `--if-unchanged` must equal the input's SHA-256 (`check` prints it) → otherwise `E-CONFLICT`. A source or output that changes during apply → `E-CONFLICT`, and the other writer's file is preserved.
- `apply` refuses on any error.
- **Residual limit.** Excel and other programs ignore advisory locks, and the final checks and `os.replace` are separate filesystem operations. A short window remains in which an uncooperative writer can still save. This is not filesystem compare-and-swap.

## 8. TEL — the Edit Language

### 8.1 Lexical structure
- Scripts are UTF-8 and line-oriented. A statement ends at the end of its line.
- `#` starts a comment only as the first non-blank character of a line, because `#` occurs inside formulas (`#REF!`).
- Keywords are case-insensitive: `in let set clear expect insert delete rows cols add rename sheet to`.
- Strings are `"…"` with `""` escapes. Numbers may be negative. Lists use `[ … ]` and may nest one level.
- The scanner is driven by the parser. After `set TARGET`, an `=` switches it to **formula mode**: the rest of the line is one FORMULA token, handed to the formula compiler with its column offset. `expect` always takes the rest of the line as a FORMULA. This is a lexer-mode "island grammar". Targets are scanned as one word (quoted sheet names may contain spaces) and parsed by the formula parser's reference rule.

### 8.2 Grammar (EBNF)
```ebnf
script     = { line } EOF ;
line       = [ statement ] NEWLINE ;
statement  = "in" sheetname "{"                       (* opens a block *)
           | "}"                                       (* closes it *)
           | "let" IDENT "=" ( target | literal )
           | "set" target ( literal | list | FORMULA )
           | "clear" target
           | "expect" FORMULA
           | ( "insert" | "delete" ) ( "rows" ROWSPAN | "cols" COLSPAN )
           | "add" "sheet" STRING
           | "rename" "sheet" STRING "to" STRING ;
sheetname  = STRING | IDENT ;
target     = [ SHEET "!" ] ( CELL [ ":" CELL ] ) | IDENT ;
literal    = NUMBER | STRING | BOOLEAN ;
list       = "[" item { "," item } "]" ;   item = literal | list ;   (* ≤ 2 levels *)
ROWSPAN    = INT [ ":" INT ] ;  COLSPAN = COL [ ":" COL ] ;
```
**Error recovery.** A syntax error in a statement is reported and parsing resumes at the next line: newline-synchronised panic mode. Unbalanced braces are reported at the stray `}` or at end of file, naming the line of the unclosed `in`.

### 8.3 Static semantics (resolver)

**Scopes.**
- The global scope has no current sheet. `in "S" { … }` pushes a scope whose current sheet is `S`; blocks nest.
- `let` declares in the current scope. Redeclaring in the same scope is `E-DECL`, shadowing an outer name is `W-SHADOW`, and a name is visible only after its declaration.
- A `let` range is a fixed address, like a defined name. It is inlined into formulas as an **absolute** reference (`$E$2:$E$21`), so a fill never moves it. A defined name used as a *target* follows the structural edits made earlier in the script, as Excel's names do.

**Sheets.**
- The resolver keeps a sheet table that `add sheet` and `rename sheet` update in program order. Unknown sheets give `E-SHEET` with the closest names by Levenshtein distance.
- An unqualified target or `expect` reference outside any `in` block is `E-NOSHEET` when the workbook has more than one sheet. Agents must say where they write.

**Targets.**
- A target is a cell, a cell range, or a name bound to one.
- Larger than 100,000 cells gives `E-SIZE`. A `let` bound to a literal used as a target gives `E-TYPE`.

**Values.**
- A literal is broadcast to every target cell.
- A list must match the target's shape: 1-D lists fit a 1×n or n×1 target, 2-D lists fit r×c (`E-SHAPE`).
- A formula is bound as in §4.1. Unqualified references inside a formula refer to the **target cell's sheet** (Excel semantics). If that differs from the scope's sheet, the resolver emits `W-SCOPE`.

**Structural statements** act on the scope's current sheet. Insert and delete take rows `1…1,048,576` or columns `A…XFD`.

### 8.4 Lowering and the edit IR
```
SetCell(sheet, col, row, content, line)       content = literal | formula AST (already filled)
ClearCell(sheet, col, row, line)
InsertRows(sheet, at, n, line)  DeleteRows(…)  InsertCols(…)  DeleteCols(…)
AddSheet(name, line)  RenameSheet(old, new, line)
Expect(host_sheet, ast, source, line)
```
Each op is a quadruple-like tuple: operator, sheet, address, operand. A formula written to a range is expanded per cell with `shift_for_fill` relative to the target's top-left cell.

### 8.5 IR optimisation — dead-write elimination
The pass scans the op list backwards with a set of cells that will be written later:
- A `SetCell` or `ClearCell` whose cell is already in the set is dead. It is removed and reported as `W-DEAD-WRITE`, aggregated per statement pair.
- `Expect` (which reads), structural ops and sheet ops are **barriers** that clear the set.

This is classic dead-store elimination, with reads and address-space changes as barriers.

### 8.6 Simulation and analysis
Simulation executes the IR on the in-memory workbook, in order:

| Op | Effect |
|---|---|
| `SetCell` / `ClearCell` | `--allow` check (`E-PERM`); constant over formula (`W-OVERWRITE-FORMULA`); replace content, keeping the `uid`; mark `touched`; update graph edges |
| Structural | Refuse (`E-STRUCT`) when the workbook has any address-bearing feature Tabula does not relocate, any unparsed or unsupported formula/name, or when an insert would push cells/formatting off the grid. Otherwise relocate every formula, defined name and pending `expect`; move the sheet's cells and formatting metadata; mark changed formulas `text_changed` |
| `AddSheet` / `RenameSheet` | Update the sheet table; `rename_sheet` over all ASTs and names. A rename is refused under the same workbook-wide rule as structural edits, and a refused rename stops the script, because later statements were resolved against the new name |
| `Expect` | Queued, and relocated by later structural ops |

After the last op:
1. **Recompute.** Incremental from the touched cells. After a structural op the graph is rebuilt, and the seeds are:
   - touched cells;
   - relocated formulas;
   - **every formula that reads an edited sheet**, because a formula over `A:A` or a relocated defined name keeps its text but sees different data.
2. **New cycles** (cycle members that were not cycle members before the edit) give `E-CYCLE` with the path.
3. **Results:**
   - *writes*: touched cells with before → after;
   - *affected*: untouched formulas whose value changed, matched by `uid`;
   - *new_errors*: written or affected cells that now hold an error and did not before (`W-NEW-ERROR`);
   - *unverified*: cells whose value is not known.
4. **Expects.** Each `expect` is evaluated on the final state with its host sheet. FALSE, an error or a non-boolean is `E-EXPECT`; a dependency on unverified values is `W-EXPECT-UNVERIFIED`.

### 8.7 Example (traceable)

On the demo workbook:

```
in "Sales" {
  insert rows 11
  set A11:D11 ["West", "Dana Lee", 40, 125]
  set E11 =C11*D11
  set F11 =E11*CommissionRate
}
expect Summary!B3 = Sales!E23
```

1. **Resolver:** `Sales` exists; the list shape 1×4 matches `A11:D11`; `CommissionRate` is a workbook defined name; `Summary` exists.
2. **IR:** `InsertRows(Sales, 11, 1)`, then six `SetCell`s, then `Expect`.
3. **Simulation:**
   - `Sales!E22 =SUM(E2:E21)` moves to `E23` and becomes `=SUM(E2:E22)` (the range spanned row 11, so it expands).
   - `Summary!B3 =Sales!E22` becomes `=Sales!E23`.
   - The defined name `SalesRevenue` (`Sales!$E$2:$E$21`) becomes `Sales!$E$2:$E$22`.
   - After recompute, E23 includes the new 5,000.
   - The expect passes.
4. **Apply:** openpyxl `insert_rows(11)`, then the 6 written cells plus every relocated formula are rewritten, then saved, verified and renamed into place.

With plain openpyxl, `insert_rows(11)` alone leaves `=SUM(E2:E21)` in E23 (which now excludes the last data row), `Summary!B3` pointing at a data row, and the defined name stale. The demo measures exactly this class of failure.

---

## 9. Error-Handling Strategy

| Layer | Detection | Containment | User / agent sees |
|---|---|---|---|
| TEL lexer and parser | bad tokens, malformed statements | skip to the next line | positioned `E-LEX` / `E-SYNTAX` |
| Formula front end | lexical and syntax errors in embedded formulas | statement rejected | positions mapped into the script |
| Resolver | sheets, names, functions, shapes, sizes, scopes | statement rejected; later statements still checked | `E-*` with hints |
| Simulation | cycles, expectations, permissions, structural limits | whole script rejected | `E-*`; plan still shows what was computed |
| Engine | runtime error values; cycles; unsimulated functions | per cell; taint for unknowns | error values; *unverified* list |
| I/O | fidelity loss, locks, concurrent changes, verification failures | input file and competing writers' files never overwritten | `E-FIDELITY` / `E-LOCKED` / `E-CONFLICT` / `E-VERIFY` |

The script is a transaction: either every op is applied or none is.

---

## 10. A/B Agent Demo (`demo/`)

**Goal:** the same agent, task and starting workbook, edited two ways, and scored by a checker that does not use Tabula's engine.

| Part | Content |
|---|---|
| `build_workbook.py` | Generates `sales_q3.xlsx`, with sheets Inputs / Sales / Summary, defined names (`TaxRate`, `CommissionRate`, `Target`, `SalesRevenue`), cross-sheet formulas and two `TEST` rows. Also generates `sales_q3_buggy.xlsx`, with two seeded Summary bugs |
| `tasks.py` | Five tasks with natural-language prompts and value-based checks: **T1** add a Tax column and a summary line · **T2** insert a missed sale after row 10 · **T3** delete the `TEST` rows · **T4** fix two wrong Summary figures · **T5** rename `Sales` to `Q3 Sales` |
| Checks | Expected numbers are computed in plain Python from the generator's data. Checks cover cell values (evaluated by **pycel**, an independent formula engine), formulas not hard-coded, a liveness probe (change an input and re-evaluate), no error values or circular references, and untouched data intact |
| `run.py` | For each task × arm × trial: a fresh run folder holding the workbook, then `claude -p` in `--safe-mode` (no user plugins, hooks, skills or MCP) with an arm-specific tool allowlist. The **default** arm gets `python3` + openpyxl. The **tabula** arm gets `./tabula` plus `TEL.md`. The prompt, model and budget are otherwise identical. Saves the transcript (stream-json), cost, duration and the output workbook |
| `evaluate.py` | Scores every run and writes `REPORT.md`: per-task pass/fail per check, success rate per arm, **silent failures** (the agent reported success but checks failed), cost and time |
| `reference/` | Deterministic solutions: TEL scripts for all tasks, plus representative openpyxl scripts. They test the checkers and allow an offline demo |

**Threats to validity, stated in the report:**
- The tabula arm receives a language reference the default arm does not need.
- Results are a small sample from one model.
- pycel's function coverage limits the checks to the functions the demo workbook uses.

---

## 11. Testing Strategy

| Level | Contents |
|---|---|
| Formula unit | Token classes incl. sheet qualifiers, `$A`, `1:3`, error literals; precedence table incl. left-assoc `^` and postfix `%`; emitter round-trip over a corpus; every supported function and coercion rule |
| Relocation | Table-driven: insert above / inside / below ranges, absolute refs, whole-column and whole-row, delete partial / whole → `#REF!`, cross-sheet, names, rename with quoting |
| Engine | Edge maintenance, range edges, closure, Kahn order, Tarjan cycles incl. cross-sheet, 5,000-deep chain without recursion, taint propagation, incremental ≡ full recompute |
| TEL | Each statement form; each diagnostic code (one positive + one negative case); line recovery; dead-write elimination; fill expansion; simulation results; JSON golden snapshot |
| I/O | Round trip; atomicity (failure injected during save leaves the input unchanged); verification incl. moved and unexpected cells; fidelity guard; snapshot load; concurrent apply writers serialise; external save during apply is preserved; Excel owner file at commit time; 1904 dates; grid overflow (`tests/test_boundary_safety.py`) |
| Dependencies | SUMIF effective ranges and implicit cycles; conservative invalidation of unsupported formulas; randomised incremental ≡ fresh full recompute (`tests/test_dependency_safety.py`) |
| Evaluation | 15 held-out tasks; reference solutions pass; deliberately corrupted workbooks (constants, styles, unrelated rows) fail (`tests/test_demo_suite.py`) |
| Differential | Tabula value vs Excel-cached value for every simulable formula in each corpus workbook saved by Excel |
| Benchmark | Generated 10k-formula workbook: incremental vs full recompute (NFR-2) |
| Demo | Reference TEL solutions pass every check; representative naive openpyxl solutions fail the checks they should |

---

## 12. Design Decisions & Trade-offs

| Decision | Alternatives | Rationale |
|---|---|---|
| A *language* for agents (TEL) | JSON tool calls (MCP), raw openpyxl | A script is a whole transaction that can be compiled and analysed *before* execution (scopes, dead writes, cycles, impact). Tool calls are checked one at a time, after the fact |
| Line-oriented TEL with formula islands | One combined grammar | Formulas are Excel's own language. Reusing the formula compiler on the rest of the line keeps one source of truth, and positions map back exactly |
| Excel-compatible formula semantics | Own dialect (Revision 1) | Tabula writes into real workbooks; any divergence would be a correctness bug |
| openpyxl at the boundary only | Hand-written OOXML writer | File format is not a compiler concern. Isolating it keeps the dependency rule clean and the compiler entirely hand-written |
| Symbolic range edges, bucketed by column | Expand ranges into cells | Whole-column references have 1,048,576 cells. A lookup scans only the ranges over that column; ranges wider than 64 columns sit in a per-sheet list. The bucketing cut a 2,500-deep chain edit from 962 ms to 10 ms (`tools/bench.py`) |
| Full rebuild after structural edits | Incremental relocation of graph edges | Structural edits are rare. One proven code path is worth more than speed here |
| Opaque cells + taint instead of guessing | Approximate unsupported functions | A verifier that guesses is worse than one that says "unverified" |
| Refuse structural edits **workbook-wide** when any address-bearing feature or unsupported formula exists | Per-sheet refusal with reference tracking; relocate those features | Rules, names and `INDIRECT` strings can point at a sheet indirectly, so per-sheet detection cannot prove safety. Workbook-wide refusal is safe, but blocks structural edits in many feature-rich workbooks; relocating these features is the main future enhancement |
| Bytecode VM moved to stretch | Keep Revision 1 Phase 3 | The edit IR, its optimiser and the back end cover IR, optimisation and code generation with real payoff |
| Headless Claude Code as the demo agent | Hand-rolled API agent | It is the agent people actually use with a shell; `--safe-mode` gives the same clean configuration for both arms |

## 13. Requirement Traceability

| FR | Section | FR | Section | FR | Section |
|---|---|---|---|---|---|
| FR-1 | §6.1 | FR-10 | §6.3 | FR-19 | §8.6 |
| FR-2 | §3 | FR-11 | §7 | FR-20 | §8.6 |
| FR-3 | §4 | FR-12 | §2.1, §6.3 | FR-21 | §5.5, §8.6 |
| FR-4 | §4.1, §5.2 | FR-13 | §7 | FR-22 | §2.1, §7 |
| FR-5 | §5.3 | FR-14 | §8.1 | FR-23 | §2.3 |
| FR-6 | §4.2, §5.4 | FR-15 | §8.2 | FR-24 | §2.1 |
| FR-7 | §6.2 | FR-16 | §8.3 | FR-25 | §10 |
| FR-8 | §6.2 | FR-17 | §8.4 | FR-26 | §9 |
| FR-9 | §6.3 | FR-18 | §8.5 | | |
