"""TEL compiler driver: source -> front end -> IR -> optimiser -> simulation -> plan,
and for `apply`, -> back end (design.md sections 2.1, 7, 8).
"""
from __future__ import annotations

from pathlib import Path

from .. import xlsx
from ..diagnostics import error, warning
from .ir import eliminate_dead_writes
from .parser import parse_script
from .plan import Plan
from .resolver import Resolver
from .simulate import Allowlist, simulate


def compile_source(source: str, workbook):
    """Front end + lowering + IR optimisation. Returns (script, live_ops, diagnostics, stats)."""
    script = parse_script(source)
    resolver = Resolver(workbook)
    ops, rdiags = resolver.run(script.statements)
    diags = script.diagnostics + rdiags
    live, dead = eliminate_dead_writes(ops)
    for (dead_line, killer_line), n in sorted(dead.items()):
        diags.append(warning("W-DEAD-WRITE", f"{n} cell(s) written here are overwritten on line "
                             f"{killer_line} before anything reads them; this write has no "
                             "effect", dead_line))
    stats = {"statements": resolver.statements, "ops": len(ops),
             "dead_writes": sum(dead.values())}
    return script, live, diags, stats


def check(book, script_path, allow_specs=(), source: str | None = None, predict_loss=True):
    """Compile and simulate; never writes. Returns (Plan, Loaded)."""
    book, script_path = Path(book), Path(script_path)
    if source is None:
        source = script_path.read_text(encoding="utf-8")
    loaded = xlsx.load(book)
    allow = Allowlist(allow_specs) if allow_specs else None
    script, ops, diags, stats = compile_source(source, loaded.model)
    plan = Plan(str(book), str(script_path), loaded.sha256, script.lines, diags, stats=stats)
    if loaded.unparsed:
        diags.append(warning("W-UNPARSED", f"the workbook has {loaded.unparsed} formula(s) "
                             "Tabula cannot parse (e.g. array formulas); cells that depend on "
                             "them may be missing from the impact report"))
    if any(d.is_error for d in diags):
        return plan, loaded
    sim = simulate(loaded.engine, ops, allow)
    diags.extend(sim.diagnostics)
    plan.sim = sim
    plan.stats = dict(stats, cells_written=len(sim.writes), **sim.stats)
    if predict_loss:
        lost = xlsx.predict_loss(loaded)
        if lost:
            diags.append(warning("W-LOSSY", "saving this workbook with Tabula would lose: "
                                 + "; ".join(lost), hint="apply will refuse unless --allow-lossy"))
    return plan, loaded


def apply(book, script_path, out=None, in_place=False, allow_specs=(), if_unchanged=None,
          allow_lossy=False) -> Plan:
    """check, then write atomically if (and only if) there are no errors."""
    book = Path(book)
    output = book if in_place else Path(out)
    try:
        with xlsx.writer_locks(book, output):
            destination_hash = xlsx.sha256(output) if output.exists() else None
            return _apply_locked(book, script_path, output, allow_specs, if_unchanged,
                                 allow_lossy, destination_hash)
    except (xlsx.WriteRefused, OSError) as exc:
        code = exc.code if isinstance(exc, xlsx.WriteRefused) else "E-WRITE"
        plan = Plan(str(book), str(script_path), "", [], [])
        plan.diagnostics.append(error(code, str(exc)))
        return plan


def _apply_locked(book, script_path, output, allow_specs, if_unchanged,
                  allow_lossy, destination_hash):
    plan, loaded = check(book, script_path, allow_specs, predict_loss=False)
    for path in {book, output}:
        lock = xlsx.lock_file(path)
        if lock is not None:
            plan.diagnostics.append(error("E-LOCKED", f"the workbook is open in Excel ({lock.name} "
                                          "exists); close it first or your edits could be lost"))
    if if_unchanged and if_unchanged.lower() != loaded.sha256:
        plan.diagnostics.append(error("E-CONFLICT", "the workbook changed since it was checked "
                                      f"(sha256 {loaded.sha256[:12]}..., expected "
                                      f"{if_unchanged[:12]}...); run check again"))
    if not plan.ok:
        return plan
    try:
        plan.synced = xlsx.write_back(loaded, plan.sim.structure, output, allow_lossy,
                                     destination_hash=destination_hash)
    except xlsx.WriteRefused as e:
        plan.diagnostics.append(error(e.code, e.message))
        return plan
    except OSError as e:
        plan.diagnostics.append(error("E-WRITE", f"could not write {output}: {e}"))
        return plan
    plan.applied, plan.output = True, str(output)
    return plan
