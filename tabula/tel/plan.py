"""Plans: the compiler's report of what a script does (design.md section 2.3).

Text for people and agents reading a terminal; JSON (schema 1) for programs.
Lists are capped but counts are always exact, so output stays small enough
for an agent's context window on large edits.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field

from ..values import BLANK, display

TEXT_LIMIT = 12
JSON_LIMIT = 50


@dataclass
class Plan:
    workbook: str
    script: str
    sha256: str
    source_lines: list
    diagnostics: list = field(default_factory=list)
    sim: object = None                   # SimResult, None when the front end failed
    stats: dict = field(default_factory=dict)
    applied: bool = False
    output: str | None = None
    synced: int = 0

    @property
    def errors(self) -> list:
        return [d for d in self.diagnostics if d.is_error]

    @property
    def ok(self) -> bool:
        return not self.errors

    def sorted_diagnostics(self) -> list:
        return sorted(self.diagnostics, key=lambda d: (d.line or 10**9, not d.is_error, d.col))


def _value(v) -> object:
    if v is BLANK:
        return None
    if isinstance(v, (bool, int, float, str)):
        return v
    return display(v)


def _change_json(ch) -> dict:
    return {"cell": ch.address,
            "before": {"raw": ch.before_raw, "value": _value(ch.before_value)},
            "after": {"raw": ch.after_raw, "value": _value(ch.after_value)},
            "unverified": ch.unverified}


def render_json(plan: Plan, limit: int = JSON_LIMIT) -> str:
    sim = plan.sim

    def capped(items, conv):
        return {"count": len(items), "items": [conv(x) for x in items[:limit]]}

    out = {
        "schema": 1,
        "ok": plan.ok,
        "applied": plan.applied,
        "workbook": plan.workbook,
        "script": plan.script,
        "sha256": plan.sha256,
        "diagnostics": [d.to_json() for d in plan.sorted_diagnostics()],
        "writes": capped(sim.writes if sim else [], _change_json),
        "affected": capped(sim.affected if sim else [], _change_json),
        "new_errors": capped(sim.new_errors if sim else [], _change_json),
        "unverified": capped(sim.unverified if sim else [], str),
        "relocated_formulas": sim.relocated if sim else 0,
        "expects": sim.expects if sim else [],
        "stats": plan.stats,
    }
    if plan.applied:
        out["output"] = plan.output
        out["cells_synced"] = plan.synced
    return json.dumps(out, indent=2, ensure_ascii=False)


def _fmt_change(ch, show_before: bool) -> str:
    after = ch.after_raw if ch.after_raw != "" else "(cleared)"
    line = f"  {ch.address:<16} {after}"
    if ch.after_raw.startswith("="):
        line += f"  -> {display(ch.after_value)}"
    if ch.unverified:
        line += "  (unverified)"
    if show_before and ch.before_raw not in ("",):
        line += f"   [was {ch.before_raw}]"
    return line


def render_text(plan: Plan, limit: int = TEXT_LIMIT) -> str:
    out = [f"tabula: {plan.script} on {plan.workbook}"]
    for d in plan.sorted_diagnostics():
        out.append(d.render(plan.source_lines, plan.script))
    sim = plan.sim
    errors = plan.errors
    if errors:
        out.append(f"\nREJECTED: {len(errors)} error(s). Nothing was written.")
    elif plan.applied:
        out.append(f"\nAPPLIED: wrote {plan.output} ({plan.synced} cells synced, re-read and "
                   "verified).")
    else:
        out.append("\nOK: no errors. Nothing has been written yet (this was a check).")
    if sim is None:
        return "\n".join(out)

    def section(title, items, fmt):
        if not items:
            return
        out.append(f"\n{title} ({len(items)})")
        out.extend(fmt(x) for x in items[:limit])
        if len(items) > limit:
            out.append(f"  ... {len(items) - limit} more")

    section("writes", sim.writes, lambda ch: _fmt_change(ch, True))
    section("affected formulas (value changes)", sim.affected,
            lambda ch: f"  {ch.address:<16} {display(ch.before_value)} -> "
                       f"{display(ch.after_value)}" + ("  (unverified)" if ch.unverified else ""))
    if sim.relocated:
        out.append(f"\nrelocated: {sim.relocated} formula(s) rewritten to follow moved/renamed cells")
    section("unverified (Tabula cannot compute these; Excel will)", sim.unverified,
            lambda a: f"  {a}")
    if sim.expects:
        out.append("\nexpects")
        for e in sim.expects:
            out.append(f"  line {e['line']:<4} {e['source']:<40} {e['status'].upper()}")
    s = plan.stats
    out.append(f"\nstats: {s.get('statements', 0)} statements -> {s.get('ops', 0)} ops "
               f"({s.get('dead_writes', 0)} dead writes removed); "
               f"{s.get('formulas_recomputed', 0)} of {s.get('total_formulas', 0)} formulas "
               f"recomputed in {s.get('recompute_ms', 0)} ms")
    out.append(f"sha256: {plan.sha256}")
    return "\n".join(out)
