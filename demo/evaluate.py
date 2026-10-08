"""Score demo runs and write REPORT.md (design.md section 10).

Run with the demo virtualenv, which has pycel (an independent Excel formula
engine) -- the checker never uses Tabula's own engine to judge Tabula:

    demo/.venv/bin/python demo/evaluate.py demo/runs/<run-folder>
"""
from __future__ import annotations

import json
import logging
import math
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from pycel import ExcelCompiler

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
logging.disable(logging.CRITICAL)  # pycel logs every evaluation failure; checks report them instead

from build_workbook import COMMISSION_RATE, DATA, TAX_RATE  # noqa: E402
from tasks import TASKS, TASKS_BY_ID, figures, rows_after  # noqa: E402

ARMS = ("default", "tabula")
ARM_LABEL = {"default": "default (Python + openpyxl)", "tabula": "Tabula (TEL)"}


# ---------------------------------------------------------------- workbook access
class Book:
    def __init__(self, path: Path):
        self.path = path
        self.wb = openpyxl.load_workbook(path)
        self._xl = None

    def compiler(self, fresh: bool = False):
        if fresh or self._xl is None:
            xl = ExcelCompiler(filename=str(self.path), cycles=False)
            if not fresh:
                self._xl = xl
            return xl
        return self._xl

    def raw(self, sheet: str, addr: str):
        if sheet not in self.wb.sheetnames:
            return None
        return self.wb[sheet][addr].value

    def value(self, addr: str, xl=None):
        try:
            return (xl or self.compiler()).evaluate(addr)
        except Exception as e:  # pycel raises on cycles, bad formulas, missing sheets
            return EvalFailure(_reason(e))


class EvalFailure(str):
    pass


def _reason(e: Exception) -> str:
    """Short, readable reason from a pycel exception."""
    text = str(e)
    if "RecursionError" in text or isinstance(e, RecursionError):
        return "circular reference"
    if "does not exist" in text and "Worksheet" in text:
        return "refers to a sheet that does not exist"
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    tail = lines[-1] if lines else type(e).__name__
    return f"{type(e).__name__}: {tail[:120]}"


def q(sheet: str) -> str:
    return f"'{sheet}'" if " " in sheet else sheet


def is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def close(a, b) -> bool:
    return is_num(a) and math.isclose(a, b, rel_tol=1e-6, abs_tol=1e-6)


def is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def show(v) -> str:
    if isinstance(v, float):
        return f"{v:,.4g}" if abs(v) < 1e4 else f"{v:,.2f}"
    return repr(v) if isinstance(v, str) else str(v)


# ---------------------------------------------------------------- checks
class Checks:
    def __init__(self):
        self.items: list[dict] = []

    def add(self, name: str, ok: bool, detail: str = "") -> None:
        self.items.append({"check": name, "ok": bool(ok), "detail": "" if ok else detail})

    def values(self, book: Book, name: str, pairs, formula_sheet: str | None = None) -> None:
        """pairs: [(address, expected)]; numbers compared approximately."""
        bad = []
        for addr, expected in pairs:
            got = book.value(addr)
            ok = close(got, expected) if is_num(expected) else (
                isinstance(got, str) and got.strip().lower() == str(expected).lower())
            if not ok:
                bad.append(f"{addr}={show(got)} (expected {show(expected)})")
        self.add(name, not bad, "; ".join(bad[:4]) + (f" (+{len(bad) - 4} more)" if len(bad) > 4 else ""))

    def formulas(self, book: Book, name: str, sheet: str, addrs) -> None:
        bad = [a for a in addrs if not is_formula(book.raw(sheet, a))]
        self.add(name, not bad, "not formulas (hard-coded or empty): " + ", ".join(bad[:6]))

    def data_rows(self, book: Book, sheet: str, data, first: int = 2) -> None:
        bad = []
        for i, row in enumerate(data):
            r = first + i
            got = tuple(book.raw(sheet, f"{c}{r}") for c in "ABCD")
            if got != row:
                bad.append(f"row {r}: {got} != {row}")
        self.add("data rows intact and in order", not bad, "; ".join(bad[:3]))

    def no_errors(self, book: Book) -> None:
        bad = []
        xl = book.compiler()
        for ws in book.wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if is_formula(c.value):
                        addr = f"{q(ws.title)}!{c.coordinate}"
                        v = book.value(addr, xl)
                        if isinstance(v, EvalFailure) or (isinstance(v, str) and v.startswith("#")):
                            bad.append(f"{addr}: {v}")
        self.add("no error values or broken/circular formulas", not bad, "; ".join(bad[:4]))

    def live(self, book: Book, name: str, set_addr: str, new_value, probes) -> None:
        """Change an input and re-evaluate: hard-coded numbers will not follow."""
        try:
            xl = book.compiler(fresh=True)
            for addr, _ in probes:
                xl.evaluate(addr)
            xl.set_value(set_addr, new_value)
            bad = []
            for addr, expected in probes:
                got = xl.evaluate(addr)
                if not close(got, expected):
                    bad.append(f"{addr}={show(got)} (expected {show(expected)})")
        except AssertionError as e:
            bad = [f"the result no longer depends on {set_addr} (the link is broken)"
                   if "cell map" in str(e) else _reason(e)]
        except Exception as e:
            bad = [_reason(e)]
        self.add(name, not bad, "; ".join(bad[:3]))

    def unchanged(self, book: Book, original: Book, sheet: str) -> None:
        a = {c.coordinate: c.value for row in original.wb[sheet].iter_rows() for c in row
             if c.value is not None}
        if sheet not in book.wb.sheetnames:
            self.add(f"sheet {sheet} unchanged", False, "sheet missing")
            return
        b = {c.coordinate: c.value for row in book.wb[sheet].iter_rows() for c in row
             if c.value is not None}
        diff = sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))
        self.add(f"sheet {sheet} unchanged", not diff, "changed: " + ", ".join(diff[:6]))


def summary_pairs(f: dict, sheet: str = "Summary") -> list:
    return [(f"{sheet}!B3", f["total_revenue"]), (f"{sheet}!B4", f["total_commission"]),
            (f"{sheet}!B5", f["average"]), (f"{sheet}!B6", f["north"]),
            (f"{sheet}!B7", f["target_met"])]


def check_task(task_id: str, path: Path) -> list[dict]:
    task = TASKS_BY_ID[task_id]
    c = Checks()
    try:
        book = Book(path)
    except Exception as e:
        c.add("workbook opens", False, f"{type(e).__name__}: {e}")
        return c.items
    original = Book(HERE / "workbooks" / task.workbook)
    data = rows_after(task_id)
    f = figures(data)
    n = len(data)
    last, total_row = 1 + n, 2 + n

    if task_id == "t1_tax_column":
        rev = f["revenue"]
        c.add("G1 header is 'Tax'", str(book.raw("Sales", "G1")).strip().lower() == "tax",
              f"G1={book.raw('Sales', 'G1')!r}")
        c.values(book, "tax per row = revenue x tax rate",
                 [(f"Sales!G{2 + i}", r * TAX_RATE) for i, r in enumerate(rev)])
        c.formulas(book, "tax cells are formulas", "Sales", [f"G{r}" for r in range(2, last + 1)])
        c.values(book, "tax total in row 22", [(f"Sales!G{total_row}", f["total_tax"])])
        label = str(book.raw("Summary", "A8") or "")
        c.add("Summary A8 label mentions tax", "tax" in label.lower(), f"A8={label!r}")
        c.values(book, "Summary B8 = total tax", [("Summary!B8", f["total_tax"])])
        c.formulas(book, "Summary B8 is a formula", "Summary", ["B8"])
        c.values(book, "existing Summary figures unchanged", summary_pairs(f))
        c.live(book, "tax follows the tax rate (live, not hard-coded)", "Inputs!B2", 0.25,
               [("Sales!G2", rev[0] * 0.25), ("Summary!B8", f["total_revenue"] * 0.25)])
    elif task_id in ("t2_insert_row", "t3_delete_test_rows"):
        if task_id == "t2_insert_row":
            c.add("new sale in row 11",
                  tuple(book.raw("Sales", f"{x}11") for x in "ABCD") == ("West", "Dana Lee", 40, 125),
                  f"row 11 = {tuple(book.raw('Sales', f'{x}11') for x in 'ABCD')}")
            c.values(book, "row 11 revenue and commission",
                     [("Sales!E11", 5000), ("Sales!F11", 5000 * COMMISSION_RATE)])
            c.formulas(book, "row 11 calculations are formulas", "Sales", ["E11", "F11"])
        else:
            tests = [r for r in range(1, 40) if book.raw("Sales", f"B{r}") == "TEST"]
            c.add("no TEST rows remain", not tests, f"TEST still in rows {tests}")
        c.data_rows(book, "Sales", data)
        c.values(book, "revenue/commission per row",
                 [(f"Sales!E{2 + i}", rv) for i, rv in enumerate(f["revenue"])])
        c.add(f"Total row is row {total_row}",
              str(book.raw("Sales", f"A{total_row}")).strip().lower() == "total",
              f"A{total_row}={book.raw('Sales', f'A{total_row}')!r}")
        c.values(book, "Sales totals include exactly the data rows",
                 [(f"Sales!C{total_row}", f["units"]), (f"Sales!E{total_row}", f["total_revenue"]),
                  (f"Sales!F{total_row}", f["total_commission"])])
        c.values(book, "every Summary figure correct", summary_pairs(f))
        c.live(book, "Summary follows a change to the first sale", "Sales!C2", 100,
               [("Summary!B3", f["total_revenue"] + (100 - data[0][2]) * data[0][3])])
    elif task_id == "t4_fix_summary":
        c.values(book, "total commission fixed (B4)", [("Summary!B4", f["total_commission"])])
        c.values(book, "North revenue fixed (B6)", [("Summary!B6", f["north"])])
        c.formulas(book, "fixes are formulas", "Summary", ["B4", "B6"])
        c.values(book, "every Summary figure correct", summary_pairs(f))
        c.unchanged(book, original, "Sales")
        c.live(book, "B4 follows the commission rate", "Inputs!B3", 0.1,
               [("Summary!B4", f["total_revenue"] * 0.1)])
        c.live(book, "B6 follows a North sale", "Sales!C2", 100,
               [("Summary!B6", f["north"] + (100 - DATA[0][2]) * DATA[0][3])])
    elif task_id == "t5_rename_sheet":
        names = book.wb.sheetnames
        c.add("sheet renamed to 'Q3 Sales'", "Q3 Sales" in names and "Sales" not in names,
              f"sheets: {names}")
        c.values(book, "every Summary figure correct", summary_pairs(f))
        dn = book.wb.defined_names.get("SalesRevenue")
        target = dn.attr_text if dn is not None else None
        c.add("named range SalesRevenue points at 'Q3 Sales'",
              target is not None and "q3 sales" in target.lower(), f"SalesRevenue -> {target!r}")
        if "Q3 Sales" in names:
            c.data_rows(book, "Q3 Sales", data)
        c.live(book, "Summary follows the renamed sheet", "'Q3 Sales'!C2", 100,
               [("Summary!B3", f["total_revenue"] + (100 - data[0][2]) * data[0][3])])
    if task_id == "t1_tax_column":
        c.data_rows(book, "Sales", data)
        # The tax addition permits only G cells and Summary row 8 to change.
        from suite_evaluate import cell_style, cells, schema
        for sheet in original.wb.sheetnames:
            allowed = (lambda a: a.startswith("G")) if sheet == "Sales" else ((lambda a: a in ("A8", "B8")) if sheet == "Summary" else (lambda a: False))
            a, b = cells(original.wb[sheet]), cells(book.wb[sheet])
            changed = [addr for addr in set(a) | set(b) if not allowed(addr) and (addr not in a or addr not in b or a[addr].value != b[addr].value or cell_style(a[addr]) != cell_style(b[addr]))]
            c.add(f"preservation outside tax edit on {sheet}", not changed, "changed: " + ", ".join(changed[:8]))
        c.add("workbook schema and defined names preserved", schema(book.wb) == schema(original.wb), "schema or names changed")
    c.no_errors(book)
    c.unchanged(book, original, "Inputs")
    return c.items


# ---------------------------------------------------------------- report
def evaluate_runs(out: Path) -> list[dict]:
    results = []
    for meta_path in sorted(out.glob("*/meta.json")):
        meta = json.loads(meta_path.read_text())
        checks = check_task(meta["task"], meta_path.parent / "book.xlsx")
        passed = sum(ch["ok"] for ch in checks)
        meta.update(checks=checks, passed=passed, total=len(checks),
                    success=passed == len(checks))
        meta["silent_failure"] = (not meta["success"] and not meta.get("is_error")
                                  and not meta.get("timed_out"))
        results.append(meta)
    (out / "results.json").write_text(json.dumps(results, indent=2))
    return results


def _avg(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


def write_report(out: Path, results: list[dict]) -> Path:
    first = results[0] if results else {}
    driver = first.get("driver", "?")
    lines = ["# Tabula vs default — A/B agent demo",
             "",
             f"- run folder: `{out.name}`  ·  generated {datetime.now():%Y-%m-%d %H:%M}",
             f"- driver: **{driver}**" + (f"  ·  model: `{first.get('model_id') or first.get('model')}`"
                                         if driver == "claude" else
                                         "  (deterministic scripted solutions, no LLM)"),
             "- checker: independent formula engine (pycel) + plain-Python expected values; "
             "Tabula's engine is not used to judge Tabula",
             ""]
    lines += ["## Summary", "",
              "| arm | tasks fully correct | checks passed | silent failures | avg est. cost (USD)* | avg time (s) | avg tool calls |",
              "|---|---|---|---|---|---|---|"]
    for arm in ARMS:
        rs = [r for r in results if r["arm"] == arm]
        if not rs:
            continue
        cost, secs, tools = (_avg([r.get(k) for r in rs]) for k in ("cost_usd", "duration_s", "tool_calls"))
        lines.append(f"| {ARM_LABEL[arm]} | {sum(r['success'] for r in rs)}/{len(rs)} | "
                     f"{sum(r['passed'] for r in rs)}/{sum(r['total'] for r in rs)} | "
                     f"{sum(r['silent_failure'] for r in rs)} | "
                     f"{'—' if cost is None else f'{cost:.3f}'} | {'—' if secs is None else f'{secs:.0f}'} | "
                     f"{'—' if tools is None else f'{tools:.1f}'} |")
    lines += ["", "*Silent failure*: the agent finished normally (no error, no timeout) but the "
              "workbook fails the checks — the failure an unchecked edit hides.",
              "", "\\* Claude Code's estimate at API list price. Runs use the claude.ai "
              "subscription login, where usage counts toward plan limits and is not billed.", ""]
    lines += ["## Per task", "", "| task | trap | " + " | ".join(ARM_LABEL[a] for a in ARMS) + " |",
              "|---|---|" + "---|" * len(ARMS)]
    for task in TASKS:
        cells = []
        for arm in ARMS:
            rs = [r for r in results if r["task"] == task.id and r["arm"] == arm]
            cells.append(" ".join(("✅" if r["success"] else "❌") + f" {r['passed']}/{r['total']}"
                                  for r in rs) or "—")
        if any(c != "—" for c in cells):
            lines.append(f"| {task.title} | {task.what_breaks} | " + " | ".join(cells) + " |")
    lines += ["", "## Runs", ""]
    for r in results:
        task = TASKS_BY_ID[r["task"]]
        status = "✅ correct" if r["success"] else ("❌ SILENT FAILURE" if r["silent_failure"] else "❌ failed")
        extra = []
        if r.get("cost_usd") is not None:
            extra.append(f"${r['cost_usd']:.3f}")
        if r.get("duration_s") is not None:
            extra.append(f"{r['duration_s']:.0f}s")
        if r.get("tool_calls") is not None:
            extra.append(f"{r['tool_calls']} tool calls")
        lines.append(f"### {task.title} — {ARM_LABEL[r['arm']]} (trial {r['trial']}): {status}"
                     + (f"  ·  {' · '.join(extra)}" if extra else ""))
        lines.append("")
        for ch in r["checks"]:
            lines.append(f"- {'✅' if ch['ok'] else '❌'} {ch['check']}"
                         + (f" — {ch['detail']}" if ch["detail"] else ""))
        if r.get("final_message"):
            msg = " ".join(r["final_message"].split())
            lines += ["", f"> agent's summary: {msg[:600]}{'…' if len(msg) > 600 else ''}"]
        lines.append(f"\nfiles: `{r['dir']}/` (book.xlsx, transcript)\n")
    lines += ["## Threats to validity", "",
              "- The Tabula arm receives a language reference (TEL.md); the default arm relies on the "
              "model's existing knowledge of openpyxl.",
              "- Small sample: one model, few trials per task; treat differences as illustrative, "
              "not statistically established.",
              "- The checker covers the functions this workbook uses; pycel and Excel may differ on "
              "functions outside that set.",
              "- Deterministic reference runs use hand-written naive openpyxl scripts; they show the "
              "failure mode, not how often a real agent hits it — the live runs measure that.", ""]
    path = out / "REPORT.md"
    path.write_text("\n".join(lines))
    return path


def write_summary(target: Path, folders: list[Path]) -> Path:
    """Headline comparison across several scored run folders (reads their results.json)."""
    lines = ["# A/B demo results: Tabula vs the default method", "",
             "Each row is one run folder. The same agent, prompt, model and budget are used in "
             "both arms; outputs are scored by an independent formula engine (pycel) against "
             "plain-Python expected values. Regenerate with "
             "`demo/.venv/bin/python demo/evaluate.py --summary demo/RESULTS.md demo/runs/<folder>...`.",
             "Costs are Claude Code's estimates at API list price; the runs use the claude.ai "
             "subscription, where usage counts toward plan limits instead of being billed.",
             "", "| runs | arm | tasks fully correct | silent failures | avg est. cost (USD) | avg time (s) |",
             "|---|---|---|---|---|---|"]
    examples = []
    for folder in folders:
        results = json.loads((folder / "results.json").read_text())
        if not results:
            continue
        first = results[0]
        label = ("scripted reference (no LLM)" if first.get("driver") == "reference"
                 else f"live agent: `{first.get('model_id') or first.get('model')}`, "
                      f"{max(r['trial'] for r in results)} trial(s)")
        for arm in ARMS:
            rs = [r for r in results if r["arm"] == arm]
            if not rs:
                continue
            cost, secs = _avg([r.get("cost_usd") for r in rs]), _avg([r.get("duration_s") for r in rs])
            lines.append(f"| [{folder.name}](runs/{folder.name}/REPORT.md) — {label} | {ARM_LABEL[arm]} | "
                         f"{sum(r['success'] for r in rs)}/{len(rs)} | {sum(r['silent_failure'] for r in rs)} | "
                         f"{'—' if cost is None else f'{cost:.3f}'} | {'—' if secs is None else f'{secs:.0f}'} |")
        for r in results:
            if r["silent_failure"] and first.get("driver") == "claude":
                failed = [c for c in r["checks"] if not c["ok"]]
                examples.append((folder.name, r, failed))
    lines += ["", "**Silent failure**: the agent finished normally and reported success, but the "
              "workbook fails the checks.", ""]
    lines += ["## Per task (tasks fully correct / runs)", "",
              "| task | trap | " + " | ".join(f"{f.name}: {a}" for f in folders for a in ARMS) + " |",
              "|---|---|" + "---|" * (2 * len(folders))]
    for task in TASKS:
        cells = []
        for folder in folders:
            results = json.loads((folder / "results.json").read_text())
            for arm in ARMS:
                rs = [r for r in results if r["task"] == task.id and r["arm"] == arm]
                cells.append(f"{sum(r['success'] for r in rs)}/{len(rs)}" if rs else "—")
        lines.append(f"| {task.title} | {task.what_breaks} | " + " | ".join(cells) + " |")
    if examples:
        lines += ["", "## What the silent failures looked like (live agents)", ""]
        for folder, r, failed in examples:
            task = TASKS_BY_ID[r["task"]]
            lines.append(f"- **{task.title}**, {ARM_LABEL[r['arm']]}, `{folder}` trial {r['trial']}: "
                         + "; ".join(f"{c['check']} ({c['detail'][:140]})" for c in failed[:2]))
    lines.append("")
    target.write_text("\n".join(lines))
    return target


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--summary":
        print(f"wrote {write_summary(Path(sys.argv[2]), [Path(f).resolve() for f in sys.argv[3:]])}")
        sys.exit(0)
    if len(sys.argv) != 2:
        sys.exit("usage: demo/.venv/bin/python demo/evaluate.py demo/runs/<folder>\n"
                 "       demo/.venv/bin/python demo/evaluate.py --summary OUT.md FOLDER...")
    out = Path(sys.argv[1]).resolve()
    res = evaluate_runs(out)
    print(f"wrote {write_report(out, res)}")
    for r in res:
        print(f"  {r['task']:<22} {r['arm']:<8} {'PASS' if r['success'] else 'FAIL'} "
              f"{r['passed']}/{r['total']}")
