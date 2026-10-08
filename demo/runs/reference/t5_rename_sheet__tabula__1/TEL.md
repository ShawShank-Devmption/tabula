# TEL — Tabula Edit Language (reference)

TEL is a small language for **editing Excel workbooks safely**. You write an edit script.
Tabula compiles it against the real workbook and tells you exactly what will change, before
anything is written:

- every reference and name is resolved, and typos get suggestions;
- every cell that will be written is listed, with its predicted value;
- every formula whose value changes as a result is listed;
- circular references and failed `expect` checks are rejected.

Only then does `apply` write the file, atomically.

## Workflow

```
./tabula inspect book.xlsx                  # sheets, headers, formula blocks, defined names
./tabula inspect book.xlsx Sales!A1:F10     # raw content + value of each cell in a range
./tabula deps book.xlsx Sales!E22           # what a cell reads, and what reads it
./tabula check book.xlsx edit.tel           # compile + simulate; NEVER writes
./tabula apply book.xlsx edit.tel --in-place    # check again, then write if there are no errors
                                            # (or: -o new.xlsx to write a copy)
```
Add `--json` to any command for machine-readable output. Exit code 0 means OK, and 1 means the
script has errors (nothing was written).

## Statements

One statement per line. A line starting with `#` is a comment.

| Statement | Meaning |
|---|---|
| `in "Sheet" {` … `}` | statements inside refer to this sheet by default (blocks can nest) |
| `set A1 42` · `set A1 "text"` · `set A1 TRUE` | write a constant |
| `set A1:C1 ["Item", "Qty", "Price"]` | write a row of values (a 1-D list fits one row or one column) |
| `set A2:B3 [[1, 2], [3, 4]]` | write a block of values (rows of the list = rows of the range) |
| `set D2 =B2*C2` | write an Excel formula |
| `set D2:D50 =B2*C2` | **fill**: the formula is written for D2 and copied down like Excel's fill handle (`B3*C3` in D3, …). `$` anchors stay fixed |
| `clear A1:B4` | empty the cells |
| `insert rows 11` · `insert rows 11:13` | insert rows; existing row 11 and below move down |
| `delete rows 8` · `delete rows 8:9` | delete rows; rows below move up |
| `insert cols D` · `delete cols D:E` | the same for columns |
| `add sheet "Name"` | new empty sheet |
| `rename sheet "Old" to "New"` | rename a sheet |
| `let name = Sales!E2:E21` · `let rate = 0.18` | a script-local name for a range or a value, usable as a target or inside formulas. Inside formulas a range name is written as an absolute reference (`$E$2:$E$21`), so filling a formula never moves it |
| `expect <condition>` | an assertion checked on the result, e.g. `expect Summary!B3 = Sales!E23`. A false expectation rejects the whole script |

## Rules worth knowing

- **Formulas are Excel formulas.** You can use functions, `$A$1`, `Sheet!A1`,
  `'Sheet With Spaces'!A1:B9`, `A:A`, and the workbook's defined names (shown by `inspect`).
  `^` is evaluated left to right as in Excel, so `2^3^2` is 64.
- **Say which sheet.** In a workbook with several sheets, a target outside an `in` block needs
  a sheet (`Summary!B8`).
- **Unqualified references inside a formula mean the sheet of the cell being written**, as
  in Excel. `set Summary!B8 =E22` reads `Summary!E22`; write `=Sales!E22` to read another sheet.
- **Structural edits keep formulas correct.** Inserting or deleting rows or columns rewrites
  every reference on every sheet, and every defined name, exactly as Excel would:
  - references below or right of the edit move;
  - a range that **spans** an inserted row grows to include it, but inserting directly *below*
    a range's last row does not grow it (same as Excel);
  - a range that loses rows shrinks;
  - a reference to a deleted cell becomes `#REF!`.

  Renaming a sheet updates all references to it.
- **Statements run in order.** After `insert rows 11`, the old row 11 is row 12 for every later
  statement. When deleting several separate rows, delete from the bottom up, or account for
  the shift.
- `expect` is checked after the whole script, on the final workbook.

## Reading `check` output

- **`error[...]`** — the script is rejected and nothing is written. Each error has a line,
  a caret and usually a hint, such as `did you mean 'Summary'?`.
- **`warning[...]`** — the script is allowed, but read it:
  - `W-OVERWRITE-FORMULA`: a constant replaced a formula;
  - `W-DEAD-WRITE`: a write is overwritten later;
  - `W-NEW-ERROR`: a cell now evaluates to `#DIV/0!` or another error;
  - `W-SCOPE`: unqualified references on another sheet;
  - `W-UNSIMULATED`: a function Tabula does not compute.
- **`writes`** — each written cell, its formula and its predicted value.
- **`affected formulas`** — existing formulas whose value changes as a consequence (`old -> new`).
- **`relocated`** — how many formulas were rewritten because rows, columns or sheets moved.
- **`unverified`** — values Tabula cannot compute (functions it does not simulate). Excel will
  compute them on open.

Common error codes:

| Code | Meaning |
|---|---|
| `E-SHEET` | unknown sheet |
| `E-NAME` | unknown name |
| `E-FUNC` | not an Excel function |
| `E-ARITY` | wrong argument count |
| `E-TYPE` | a range used where one value is needed |
| `E-SHAPE` | list and target sizes differ |
| `E-NOSHEET` | say which sheet |
| `E-CYCLE` | circular reference |
| `E-EXPECT` | an expectation failed |
| `E-STRUCT` | insert/delete refused because some addresses could not be updated: merged cells, tables, conditional formatting, data validation, hyperlinks, a print area or an autofilter on the sheet, or rules on another sheet that point at it |

## Examples (inventory workbook with sheets `Stock` and `Report`)

```
# add a computed column with a total, and show it on the report
in "Stock" {
  set F1 "Value"
  set F2:F40 =D2*E2
  set F41 =SUM(F2:F40)
}
set Report!A6 "Stock value"
set Report!B6 =Stock!F41
expect Report!B6 = SUM(Stock!F2:F40)
```

```
# add an item in the middle of the table, copying the row's formula pattern
in "Stock" {
  insert rows 7
  set A7:E7 ["Bolts M6", "Hardware", 120, 0.15, 0]
  set F7 =D7*E7
}
```

```
# remove two discontinued items (bottom row first, so row numbers don't shift)
in "Stock" {
  delete rows 30
  delete rows 12
}
```

```
# rename a sheet: every formula and defined name that refers to it is updated
rename sheet "Stock" to "Stock 2026"
```
