"""Tabula command line (design.md section 2.1).

    python3 -m tabula inspect BOOK.xlsx [RANGE] [--json]
    python3 -m tabula deps    BOOK.xlsx CELL [--json]
    python3 -m tabula explain BOOK.xlsx CELL
    python3 -m tabula check   BOOK.xlsx SCRIPT.tel [--json] [--allow SPEC]...
    python3 -m tabula apply   BOOK.xlsx SCRIPT.tel (-o OUT | --in-place) [--json] [--allow SPEC]...
                              [--if-unchanged SHA256] [--allow-lossy]
    python3 -m tabula repl    [SCRIPT.txt]

Exit codes: 0 success, 1 the script has errors / apply refused, 3 usage or I/O error.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import parser as P
from .refs import index_to_col, quote_sheet
from .values import display

EXIT_OK, EXIT_ERRORS, EXIT_USAGE = 0, 1, 3


class UsageError(Exception):
    pass


# ---------------------------------------------------------------- REPL (Phase 1)
HELP = """commands:
  set <cell> <literal | =formula>   edit a cell
  get <cell>                        value + raw content
  show                              render the grid
  explain <cell>                    tokens, AST, value for a cell
  help | quit"""


def run_command(sheet, line: str) -> str | None:
    from .refs import CellRef
    parts = line.strip().split(None, 2)
    if not parts:
        return None
    cmd = parts[0].lower()
    if cmd == "quit":
        return "QUIT"
    if cmd == "help":
        return HELP
    if cmd == "set" and len(parts) == 3:
        return sheet.set(parts[1], parts[2])
    if cmd == "get" and len(parts) >= 2:
        r = CellRef.parse(parts[1])
        cell = sheet.cells.get(r.key()) if r else None
        raw = f"   ({cell.raw})" if cell and cell.kind != "literal" else ""
        return f"{parts[1].upper()} = {display(sheet.value(parts[1]))}{raw}"
    if cmd == "show":
        return sheet.show()
    if cmd == "explain" and len(parts) >= 2:
        return sheet.explain(parts[1])
    return f"unrecognized command: {line.strip()!r} (try 'help')"


def repl(script: str | None) -> int:
    from .engine import Sheet
    sheet = Sheet()
    if script:
        with open(script) as f:
            for line in f:
                line = line.rstrip("\n")
                if not line.strip() or line.lstrip().startswith("#"):
                    continue
                print(f"tabula> {line}")
                out = run_command(sheet, line)
                if out == "QUIT":
                    break
                if out:
                    print(out)
        return EXIT_OK
    print("Tabula REPL - type 'help' for commands")
    while True:
        try:
            line = input("tabula> ")
        except EOFError:
            break
        out = run_command(sheet, line)
        if out == "QUIT":
            break
        if out:
            print(out)
    return EXIT_OK


# ---------------------------------------------------------------- workbook commands
def _load(path: str):
    from . import xlsx
    p = Path(path)
    if not p.exists():
        raise UsageError(f"no such file: {path}")
    try:
        return xlsx.load(p)
    except Exception as e:  # openpyxl raises many types for unreadable files
        raise UsageError(f"cannot read {path} as an .xlsx workbook: {e}")


def _resolve_cell_arg(loaded, text: str):
    """'Sales!E22' -> (sheet, node). A bare 'E22' only works for single-sheet books."""
    try:
        node = P.parse_target(text)
    except Exception as e:
        raise UsageError(f"bad cell/range {text!r}: {getattr(e, 'message', e)}")
    wb = loaded.model
    if isinstance(node, P.Name) and node.sheet is None and wb.sheet(node.ident):
        return wb.sheet(node.ident), None           # a whole sheet
    if isinstance(node, P.Name):
        dn = wb.lookup_name(node.ident, None, node.sheet)
        if dn is None or not isinstance(dn.ast, (P.Ref, P.RangeNode)):
            raise UsageError(f"unknown name or sheet {text!r}")
        node = dn.ast
    if node.sheet is None:
        if len(wb.sheets) != 1:
            raise UsageError(f"say which sheet: e.g. {quote_sheet(wb.sheets[0].name)}!{text}")
        return wb.sheets[0], node
    sheet = wb.sheet(node.sheet)
    if sheet is None:
        raise UsageError(f"unknown sheet {node.sheet!r}; sheets: "
                         + ", ".join(s.name for s in wb.sheets))
    return sheet, node


def _signature(node, col: int, row: int) -> str:
    """Formula in relative (R1C1-like) form, so filled formulas share a signature."""
    def ref(r):
        c = f"C{r.col}" if r.abs_col else f"C[{r.col - col}]"
        w = f"R{r.row}" if r.abs_row else f"R[{r.row - row}]"
        return w + c
    if isinstance(node, P.Ref):
        return f"{node.sheet}!{ref(node.ref)}"
    if isinstance(node, P.RangeNode):
        if node.rng.kind != "cells":
            return f"{node.sheet}!{node.rng}"
        return f"{node.sheet}!{ref(node.rng.start)}:{ref(node.rng.end)}"
    if isinstance(node, (P.Unary, P.Percent)):
        return f"{type(node).__name__}{getattr(node, 'op', '')}({_signature(node.operand, col, row)})"
    if isinstance(node, P.Binary):
        return f"({_signature(node.left, col, row)}{node.op}{_signature(node.right, col, row)})"
    if isinstance(node, P.Call):
        return node.fname + "(" + ",".join(_signature(a, col, row) for a in node.args) + ")"
    return repr(node)


def _formula_blocks(sheet) -> list[dict]:
    by_col: dict[int, list] = {}
    for (c, r), cell in sheet.cells.items():
        if cell.kind in ("formula", "unparsed"):
            by_col.setdefault(c, []).append((r, cell))
    blocks = []
    for c in sorted(by_col):
        run = []
        for r, cell in sorted(by_col[c], key=lambda t: t[0]):
            sig = _signature(cell.ast, c, r) if cell.kind == "formula" else cell.raw
            if run and run[-1][0] == r - 1 and run[-1][2] == sig:
                run.append((r, cell, sig))
                continue
            if run:
                blocks.append(_block(sheet, c, run))
            run = [(r, cell, sig)]
        if run:
            blocks.append(_block(sheet, c, run))
    return sorted(blocks, key=lambda b: (b["row"], b["col"]))


def _block(sheet, col, run) -> dict:
    first, last = run[0][0], run[-1][0]
    letters = index_to_col(col)
    rng = f"{letters}{first}" + (f":{letters}{last}" if last != first else "")
    cell = run[0][1]
    return {"range": rng, "formula": cell.raw, "cells": len(run), "row": first, "col": col,
            "simulated": cell.kind == "formula" and cell.simulable}


def _used_range(sheet) -> str:
    if not sheet.cells:
        return "(empty)"
    cols = [c for c, _ in sheet.cells]
    rows = [r for _, r in sheet.cells]
    return f"{index_to_col(min(cols))}{min(rows)}:{index_to_col(max(cols))}{max(rows)}"


def cmd_inspect(args) -> int:
    loaded = _load(args.book)
    wb, engine = loaded.model, loaded.engine
    if args.range:
        return _inspect_range(loaded, args)
    matching, compared, mismatches = engine.agreement()
    summary = {"workbook": args.book, "sha256": loaded.sha256, "sheets": [], "names": [],
               "engine_agreement": {"matching": matching, "compared": compared,
                                    "mismatches": [{"cell": f"{quote_sheet(s)}!{index_to_col(c)}{r}",
                                                    "tabula": display(v), "excel": display(x)}
                                                   for s, c, r, v, x in mismatches[:10]]},
               "unparsed_formulas": loaded.unparsed}
    for sheet in wb.sheets:
        header_row = min((r for _, r in sheet.cells), default=None)
        header = []
        if header_row is not None:
            header = [{"cell": f"{index_to_col(c)}{header_row}", "value": display(cell.value)}
                      for (c, r), cell in sorted(sheet.cells.items()) if r == header_row][:30]
        n_formulas = sum(1 for c in sheet.cells.values() if c.kind != "literal")
        summary["sheets"].append({
            "name": sheet.name, "used_range": _used_range(sheet), "cells": len(sheet.cells),
            "formulas": n_formulas, "first_row": header, "formula_blocks": _formula_blocks(sheet),
            "structural_blockers": sheet.structural_blockers()})
    for (scope, _), dn in sorted(wb.names.items(), key=lambda kv: kv[1].name.upper()):
        if dn.name.startswith("_xlnm."):
            continue
        summary["names"].append({"name": dn.name, "refers_to": dn.text,
                                 "scope": wb.by_sid(scope).name if scope else "workbook"})
    if args.json:
        for s in summary["sheets"]:
            for b in s["formula_blocks"]:
                b.pop("row"), b.pop("col")
        print(json.dumps(summary, indent=2))
        return EXIT_OK
    print(f"{args.book}: {len(wb.sheets)} sheet(s), sha256 {loaded.sha256[:16]}...")
    if compared:
        print(f"engine agreement: {matching}/{compared} simulated formulas match Excel's cached "
              "values" + ("" if matching == compared else " (mismatches: "
                          + ", ".join(m['cell'] for m in summary['engine_agreement']['mismatches'])
                          + ")"))
    else:
        print("engine agreement: no Excel-cached values in this file (not yet opened in Excel)")
    if loaded.unparsed:
        print(f"unparsed formulas: {loaded.unparsed} (opaque to Tabula)")
    for s in summary["sheets"]:
        print(f"\nsheet {quote_sheet(s['name'])}  used {s['used_range']}  "
              f"({s['cells']} cells, {s['formulas']} formulas)")
        if s["first_row"]:
            print("  first row: " + " | ".join(f"{h['cell']} {h['value']}" for h in s["first_row"]))
        blocks = s["formula_blocks"]
        for b in blocks[:40]:
            note = "" if b["simulated"] else "  [not simulated]"
            count = f"  ({b['cells']} cells, filled)" if b["cells"] > 1 else ""
            print(f"  {b['range']:<12} {b['formula']}{count}{note}")
        if len(blocks) > 40:
            print(f"  ... {len(blocks) - 40} more formula blocks")
        if s["structural_blockers"]:
            print("  note: insert/delete rows/cols refused here (" + ", ".join(s["structural_blockers"]) + ")")
    if summary["names"]:
        print("\ndefined names:")
        for n in summary["names"]:
            scope = "" if n["scope"] == "workbook" else f"  (sheet {n['scope']})"
            print(f"  {n['name']} = {n['refers_to']}{scope}")
    return EXIT_OK


def _inspect_range(loaded, args) -> int:
    sheet, node = _resolve_cell_arg(loaded, args.range)
    if node is None:
        cells = sorted(sheet.cells.items(), key=lambda kv: (kv[0][1], kv[0][0]))
    else:
        b = ((node.ref.col, node.ref.row, node.ref.col, node.ref.row)
             if isinstance(node, P.Ref) else node.rng.bounds())
        cells = sorted(((k, v) for k, v in sheet.cells.items()
                        if b[0] <= k[0] <= b[2] and b[1] <= k[1] <= b[3]),
                       key=lambda kv: (kv[0][1], kv[0][0]))
    limit = 400
    rows = [{"cell": f"{index_to_col(c)}{r}", "raw": cell.raw, "value": display(cell.value),
             "kind": cell.kind} for (c, r), cell in cells[:limit]]
    if args.json:
        print(json.dumps({"sheet": sheet.name, "count": len(cells), "cells": rows}, indent=2))
        return EXIT_OK
    print(f"sheet {quote_sheet(sheet.name)}: {len(cells)} non-empty cell(s)")
    for row in rows:
        if row["kind"] == "literal":
            print(f"  {row['cell']:<8} {row['raw']}")
        else:
            print(f"  {row['cell']:<8} {row['raw']}  -> {row['value']}")
    if len(cells) > limit:
        print(f"  ... {len(cells) - limit} more (narrow the range)")
    return EXIT_OK


def cmd_deps(args) -> int:
    loaded = _load(args.book)
    sheet, node = _resolve_cell_arg(loaded, args.cell)
    if not isinstance(node, P.Ref):
        raise UsageError("deps takes a single cell, e.g. Sales!E22")
    engine, wb = loaded.engine, loaded.model
    key = (sheet.sid, node.ref.col, node.ref.row)
    cell = sheet.cells.get((node.ref.col, node.ref.row))
    reads = []
    if cell is not None and cell.kind == "formula":
        cells, ranges = wb.precedents(sheet, cell.ast)
        reads = sorted(engine.address(k) for k in cells)
        reads += sorted(f"{engine.address((sid, b[0], b[1]))}:{index_to_col(b[2])}{b[3]}"
                        for sid, b in ranges)
    direct = sorted(engine.address(k) for k in engine.graph.dependents(key))
    transitive = sorted(engine.address(k) for k in engine.graph.closure({key}) - {key})
    out = {"cell": engine.address(key), "raw": cell.raw if cell else "",
           "value": display(cell.value) if cell else "", "reads": reads,
           "read_by": direct, "affects_transitively": transitive}
    if args.json:
        print(json.dumps(out, indent=2))
        return EXIT_OK
    print(f"{out['cell']}  {out['raw']}" + (f"  -> {out['value']}" if out["raw"].startswith("=") else ""))
    print("  reads   : " + (", ".join(reads) if reads else "(nothing)"))
    print("  read by : " + (", ".join(direct) if direct else "(nothing)"))
    print(f"  an edit here recomputes {len(transitive)} formula(s)"
          + (": " + ", ".join(transitive[:20]) + (" ..." if len(transitive) > 20 else "")
             if transitive else ""))
    return EXIT_OK


def cmd_explain(args) -> int:
    from .engine import explain_cell
    loaded = _load(args.book)
    sheet, node = _resolve_cell_arg(loaded, args.cell)
    if not isinstance(node, P.Ref):
        raise UsageError("explain takes a single cell, e.g. Sales!E22")
    cell = sheet.cells.get((node.ref.col, node.ref.row))
    name = loaded.engine.address((sheet.sid, node.ref.col, node.ref.row))
    if cell is None:
        print(f"{name}: (blank)")
        return EXIT_OK
    print("\n".join(explain_cell(name, cell, loaded.engine, sheet)))
    return EXIT_OK


def _plan_output(plan, as_json: bool) -> int:
    from .tel.plan import render_json, render_text
    print(render_json(plan) if as_json else render_text(plan))
    return EXIT_OK if plan.ok else EXIT_ERRORS


def cmd_check(args) -> int:
    from .tel import compiler
    _require(args.book, args.script)
    try:
        plan, _ = compiler.check(args.book, args.script, args.allow or ())
    except ValueError as e:
        raise UsageError(str(e))
    return _plan_output(plan, args.json)


def cmd_apply(args) -> int:
    from .tel import compiler
    _require(args.book, args.script)
    if bool(args.output) == bool(args.in_place):
        raise UsageError("apply needs exactly one of -o OUT.xlsx or --in-place")
    try:
        plan = compiler.apply(args.book, args.script, out=args.output, in_place=args.in_place,
                              allow_specs=args.allow or (), if_unchanged=args.if_unchanged,
                              allow_lossy=args.allow_lossy)
    except ValueError as e:
        raise UsageError(str(e))
    return _plan_output(plan, args.json)


def _require(*paths) -> None:
    for p in paths:
        if not Path(p).exists():
            raise UsageError(f"no such file: {p}")


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="tabula", description="Compile, check and apply "
                                 "spreadsheet edits (TEL) against .xlsx workbooks.")
    sub = ap.add_subparsers(dest="command", required=True)

    p = sub.add_parser("inspect", help="summarise a workbook, or list a range")
    p.add_argument("book")
    p.add_argument("range", nargs="?", help="e.g. Sales!A1:F10, or a sheet name")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("deps", help="what a cell reads and what reads it")
    p.add_argument("book")
    p.add_argument("cell")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_deps)

    p = sub.add_parser("explain", help="formula compiler pipeline for one cell")
    p.add_argument("book")
    p.add_argument("cell")
    p.set_defaults(func=cmd_explain)

    for name, fn, helptext in (("check", cmd_check, "compile and simulate a TEL script (never writes)"),
                               ("apply", cmd_apply, "check, then write the workbook if there are no errors")):
        p = sub.add_parser(name, help=helptext)
        p.add_argument("book")
        p.add_argument("script")
        p.add_argument("--json", action="store_true")
        p.add_argument("--allow", action="append", metavar="SPEC",
                       help="restrict writes: '*', Sheet, or Sheet!A1:F40 (repeatable)")
        if name == "apply":
            p.add_argument("-o", "--output", help="write the result here")
            p.add_argument("--in-place", action="store_true", help="overwrite BOOK")
            p.add_argument("--if-unchanged", metavar="SHA256",
                           help="refuse if BOOK's sha256 differs (from a previous check)")
            p.add_argument("--allow-lossy", action="store_true",
                           help="write even if openpyxl would drop workbook parts")
        p.set_defaults(func=fn)

    p = sub.add_parser("repl", help="interactive single-sheet REPL (Phase 1)")
    p.add_argument("script", nargs="?")
    p.set_defaults(func=lambda a: repl(a.script))
    return ap


COMMANDS = {"inspect", "deps", "explain", "check", "apply", "repl", "-h", "--help"}


def main(argv=None) -> None:
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] not in COMMANDS and argv[0].endswith(".txt"):
        argv = ["repl"] + argv          # Phase 1 usage: python3 -m tabula.cli demo.txt
    if not argv:
        argv = ["repl"]
    args = build_parser().parse_args(argv)
    try:
        code = args.func(args)
    except UsageError as e:
        print(f"tabula: {e}", file=sys.stderr)
        code = EXIT_USAGE
    except OSError as e:
        print(f"tabula: {e}", file=sys.stderr)
        code = EXIT_USAGE
    sys.exit(code)


if __name__ == "__main__":
    main()
