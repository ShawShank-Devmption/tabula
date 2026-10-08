"""Run the A/B demo: each task x arm x trial in its own folder, then score them.

    python3 demo/run.py                         # live: headless Claude Code, all tasks, both arms
    python3 demo/run.py --tasks t2_insert_row --trials 3 --model sonnet
    python3 demo/run.py --driver reference      # no LLM: scripted reference solutions

Arms
  default  the agent edits with Python + openpyxl (the usual way agents edit .xlsx)
  tabula   the agent edits with the Tabula CLI (./tabula) and TEL

Both arms get the same task prompt, model, budget and a clean Claude Code
configuration (--safe-mode: no user plugins, hooks, skills, CLAUDE.md or MCP).
Only the tool allowlist and a two-sentence method note differ.

Runs use your Claude subscription (the claude.ai login of the `claude` CLI),
never an API key: API-key variables are removed from the agent's environment
and the run refuses to start unless `claude auth status` reports a claude.ai
login.  Reported costs are Claude Code's estimate at API list price; on a
subscription they count toward your plan's usage limits and are not billed.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from tasks import TASKS, TASKS_BY_ID  # noqa: E402

COMMON = """You are working in a folder that contains an Excel workbook, book.xlsx.

Task: {task}

Save your changes to book.xlsx in place (keep the file name). Keep calculated values as live \
Excel formulas rather than typing in computed numbers. Only a few shell commands are available \
here; if one is refused, do the step another way (for example, create files with the Write \
tool). When you are done, reply with a short summary of what you changed."""

ARM_NOTES = {
    "default": "\n\nMethod: use Python 3 with the openpyxl library (already installed) to read and "
               "edit the workbook. Run your scripts with python3.",
    "tabula": "\n\nMethod: use the Tabula command-line tool in this folder (./tabula) to read and "
              "edit the workbook. Its language reference is TEL.md; read it first. Workflow: "
              "./tabula inspect book.xlsx, write an edit script (for example edit.tel), "
              "./tabula check book.xlsx edit.tel, fix anything it reports, then "
              "./tabula apply book.xlsx edit.tel --in-place.",
}
# identical harmless utilities for both arms; the arms differ only in python3 vs ./tabula
_COMMON_TOOLS = ["Bash(ls *)", "Bash(cat *)", "Bash(echo *)", "Bash(head *)", "Bash(tail *)",
                 "Bash(grep *)", "Bash(wc *)", "Bash(cp *)", "Read", "Write", "Edit", "Glob", "Grep"]
ALLOWED = {
    "default": ["Bash(python3 *)", "Bash(python *)"] + _COMMON_TOOLS,
    "tabula": ["Bash(./tabula *)"] + _COMMON_TOOLS,
}
TOOLS = "Bash,Read,Write,Edit,Glob,Grep"
# removed from the agent's environment so the CLI can only use the subscription login
API_ENV_VARS = ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL",
                "CLAUDE_CODE_USE_BEDROCK", "CLAUDE_CODE_USE_VERTEX", "CLAUDE_CODE_USE_FOUNDRY")


def subscription_env() -> dict:
    return {k: v for k, v in os.environ.items() if k not in API_ENV_VARS}


def require_subscription() -> dict:
    """Stop unless the claude CLI is signed in with a claude.ai (subscription) account."""
    try:
        out = subprocess.run(["claude", "auth", "status"], capture_output=True, text=True,
                             env=subscription_env(), timeout=60).stdout
        status = json.loads(out)
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError) as e:
        sys.exit(f"cannot read `claude auth status` ({e}); is Claude Code installed?")
    if not status.get("loggedIn") or status.get("authMethod") != "claude.ai":
        sys.exit("the claude CLI is not signed in with a Claude subscription; run "
                 "`claude auth login` (choose your claude.ai account) and try again")
    return status


def prepare(rundir: Path, task, arm: str) -> None:
    rundir.mkdir(parents=True)
    shutil.copy(HERE / "workbooks" / task.workbook, rundir / "book.xlsx")
    if arm == "tabula":
        shutil.copy(ROOT / "docs" / "TEL.md", rundir / "TEL.md")
        wrapper = rundir / "tabula"
        wrapper.write_text(f'#!/bin/sh\nPYTHONPATH="{ROOT}" exec "{sys.executable}" -m tabula "$@"\n')
        wrapper.chmod(0o755)


def run_claude(rundir: Path, task, arm: str, args) -> dict:
    prompt = COMMON.format(task=task.prompt) + ARM_NOTES[arm]
    cmd = ["claude", "-p", prompt, "--safe-mode", "--strict-mcp-config",
           "--no-session-persistence", "--model", args.model,
           "--output-format", "stream-json", "--verbose",
           "--max-budget-usd", str(args.budget),
           "--permission-mode", "dontAsk", "--permission-prompts", "none",
           "--tools", TOOLS, "--allowedTools", *ALLOWED[arm]]
    (rundir / "prompt.txt").write_text(prompt)
    start = time.time()
    timed_out = False
    with open(rundir / "transcript.jsonl", "w") as out, open(rundir / "stderr.txt", "w") as err:
        try:
            proc = subprocess.run(cmd, cwd=rundir, stdout=out, stderr=err, timeout=args.timeout,
                                  stdin=subprocess.DEVNULL, env=subscription_env())
            code = proc.returncode
        except subprocess.TimeoutExpired:
            timed_out, code = True, None
    meta = {"returncode": code, "timed_out": timed_out,
            "duration_s": round(time.time() - start, 1)}
    tool_calls, result = 0, None
    for line in (rundir / "transcript.jsonl").read_text().splitlines():
        try:
            ev = json.loads(line)
        except json.JSONDecodeError:
            continue
        if ev.get("type") == "assistant":
            content = ev.get("message", {}).get("content", [])
            tool_calls += sum(1 for c in content if isinstance(c, dict) and c.get("type") == "tool_use")
        elif ev.get("type") == "result":
            result = ev
    meta["tool_calls"] = tool_calls
    if result:
        meta.update(cost_usd=result.get("total_cost_usd"), num_turns=result.get("num_turns"),
                    is_error=result.get("is_error"), final_message=result.get("result") or "",
                    permission_denials=len(result.get("permission_denials") or []),
                    model_id=",".join(sorted(result.get("modelUsage") or {})))
        if result.get("duration_ms"):
            meta["duration_s"] = round(result["duration_ms"] / 1000, 1)
    else:
        meta.update(is_error=True, final_message="(no result event; see stderr.txt)")
    return meta


def run_reference(rundir: Path, task, arm: str, args) -> dict:
    env = dict(os.environ, PYTHONPATH=str(ROOT))
    if arm == "tabula":
        cmd = [sys.executable, "-m", "tabula", "apply", "book.xlsx",
               str(HERE / "reference" / f"{task.id}.tel"), "--in-place"]
    else:
        cmd = [sys.executable, str(HERE / "reference" / f"{task.id}_openpyxl.py")]
    start = time.time()
    proc = subprocess.run(cmd, cwd=rundir, env=env, capture_output=True, text=True)
    (rundir / "transcript.txt").write_text(proc.stdout + proc.stderr)
    return {"returncode": proc.returncode, "timed_out": False, "is_error": proc.returncode != 0,
            "duration_s": round(time.time() - start, 2),
            "final_message": proc.stdout.strip().splitlines()[-1] if proc.stdout.strip() else ""}


def one(job) -> dict:
    out, task, arm, trial, args = job
    rundir = out / f"{task.id}__{arm}__{trial}"
    prepare(rundir, task, arm)
    driver = run_claude if args.driver == "claude" else run_reference
    meta = {"task": task.id, "arm": arm, "trial": trial, "driver": args.driver,
            "model": args.model if args.driver == "claude" else None,
            "auth": args.auth, "dir": str(rundir.relative_to(ROOT))}
    try:
        meta.update(driver(rundir, task, arm, args))
    except Exception as e:  # keep going; the evaluator scores whatever is on disk
        meta.update(is_error=True, final_message=f"runner error: {e}")
    (rundir / "meta.json").write_text(json.dumps(meta, indent=2))
    print(f"  done {task.id:<22} {arm:<8} trial {trial}  "
          f"{meta.get('duration_s', '?')}s  cost={meta.get('cost_usd')}", flush=True)
    return meta


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--driver", choices=["claude", "reference"], default="claude")
    ap.add_argument("--tasks", default=",".join(t.id for t in TASKS))
    ap.add_argument("--arms", default="default,tabula")
    ap.add_argument("--trials", type=int, default=1)
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--budget", type=float, default=2.0, help="max USD per run")
    ap.add_argument("--timeout", type=int, default=1200, help="seconds per run")
    ap.add_argument("--parallel", type=int, default=3)
    ap.add_argument("--out", help="run folder (default demo/runs/<timestamp>-<driver>)")
    args = ap.parse_args()
    args.auth = None
    if args.driver == "claude":
        status = require_subscription()
        args.auth = f"claude.ai {status.get('subscriptionType', '')} subscription".replace("  ", " ")
        print(f"using {args.auth} ({status.get('email', '')}); API keys are ignored")
    tasks = [TASKS_BY_ID[t] for t in args.tasks.split(",")]
    arms = args.arms.split(",")
    out = (Path(args.out) if args.out else
           HERE / "runs" / f"{datetime.now():%Y%m%d-%H%M%S}-{args.driver}").resolve()
    out.mkdir(parents=True, exist_ok=True)
    jobs = [(out, t, a, k, args) for t in tasks for k in range(1, args.trials + 1) for a in arms]
    print(f"{len(jobs)} run(s) -> {out}")
    with ThreadPoolExecutor(max_workers=max(1, args.parallel)) as pool:
        list(pool.map(one, jobs))
    venv_python = HERE / ".venv" / "bin" / "python"
    if not venv_python.exists():
        print("demo/.venv missing: python3 -m venv demo/.venv && "
              "demo/.venv/bin/pip install pycel openpyxl==3.0.10   (then run evaluate.py)")
        return
    subprocess.run([str(venv_python), str(HERE / "evaluate.py"), str(out)], check=False)


if __name__ == "__main__":
    main()
