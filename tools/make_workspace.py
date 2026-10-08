"""Create a folder for trying Tabula yourself with Claude Code on your Claude subscription.

    python3 tools/make_workspace.py                       # demo workbook  -> playground/tabula
    python3 tools/make_workspace.py path/to/your.xlsx     # your own workbook
    python3 tools/make_workspace.py --method openpyxl     # comparison: plain Python + openpyxl
    cd playground/tabula && claude                        # then ask for edits in plain English

The folder gets a copy of the workbook (the original is never touched), a
CLAUDE.md telling Claude how to edit it, and a .claude/settings.json that
pre-approves exactly the commands that method needs.  Claude Code started in
the folder uses whatever you are signed in with (`claude auth status`); this
script warns if that is not a claude.ai subscription login.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEMO_BOOK = ROOT / "demo" / "workbooks" / "sales_q3.xlsx"

TABULA_GUIDE = """\
# Editing the workbook in this folder: use Tabula

This folder contains an Excel workbook ({book}) and the Tabula command-line tool (`./tabula`).
Make every change to the workbook through Tabula. Do not edit the .xlsx with Python, openpyxl
or other tools.

1. Look first.
   - `./tabula inspect {book}` shows the sheets, headers, formula blocks and defined names.
   - `./tabula inspect {book} 'Sheet!A1:F20'` shows the cells of a range.
   - `./tabula deps {book} Sheet!E22` shows what a cell reads and what reads it.
2. Write the change as a TEL script, e.g. `edit.tel`. The language reference is `TEL.md`;
   read it before your first script.
3. Run `./tabula check {book} edit.tel`. Fix every `error[...]` it reports and read the
   warnings. Add `expect` lines for what must be true afterwards.
4. Run `./tabula apply {book} edit.tel --in-place`.
5. Tell the user what changed, from the plan output: writes, affected formulas, expects.

If the user asks for something Tabula refuses (`E-STRUCT`, `E-FIDELITY`, ...), explain why
instead of working around it.
"""

OPENPYXL_GUIDE = """\
# Editing the workbook in this folder

This folder contains an Excel workbook ({book}). Use Python 3 with the openpyxl library
(already installed) to read and edit it; run scripts with `python3`.
"""


def subscription_status() -> dict | None:
    env = {k: v for k, v in os.environ.items() if k not in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")}
    try:
        out = subprocess.run(["claude", "auth", "status"], capture_output=True, text=True,
                             env=env, timeout=60).stdout
        return json.loads(out)
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError):
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", nargs="?", default=str(DEMO_BOOK), help="an .xlsx file (default: demo)")
    ap.add_argument("--method", choices=["tabula", "openpyxl"], default="tabula")
    ap.add_argument("--dir", help="workspace folder (default: playground/<method>)")
    ap.add_argument("--force", action="store_true", help="replace an existing workspace")
    args = ap.parse_args()

    src = Path(args.workbook).expanduser().resolve()
    if not src.exists() or src.suffix.lower() not in (".xlsx", ".xlsm"):
        sys.exit(f"not an .xlsx/.xlsm file: {src}")
    ws = Path(args.dir).resolve() if args.dir else ROOT / "playground" / args.method
    if ws.exists():
        if not args.force:
            sys.exit(f"{ws} already exists (use --force to replace it)")
        shutil.rmtree(ws)
    (ws / ".claude").mkdir(parents=True)
    (ws / "original").mkdir()
    book = src.name
    shutil.copy(src, ws / book)
    shutil.copy(src, ws / "original" / book)   # pristine copy, to compare or reset

    common = ["Bash(ls *)", "Bash(cat *)", "Bash(cp *)", "Read", "Write", "Edit", "Glob", "Grep"]
    if args.method == "tabula":
        wrapper = ws / "tabula"
        wrapper.write_text(f'#!/bin/sh\nPYTHONPATH="{ROOT}" exec "{sys.executable}" -m tabula "$@"\n')
        wrapper.chmod(0o755)
        shutil.copy(ROOT / "docs" / "TEL.md", ws / "TEL.md")
        (ws / "CLAUDE.md").write_text(TABULA_GUIDE.format(book=book))
        permissions = {"allow": ["Bash(./tabula *)"] + common,
                       "deny": ["Bash(python3 *)", "Bash(python *)"]}
    else:
        (ws / "CLAUDE.md").write_text(OPENPYXL_GUIDE.format(book=book))
        permissions = {"allow": ["Bash(python3 *)", "Bash(python *)"] + common}
    (ws / ".claude" / "settings.json").write_text(json.dumps({"permissions": permissions}, indent=2))

    if src == DEMO_BOOK:
        sys.path.insert(0, str(ROOT / "demo"))
        from tasks import TASKS
        lines = ["# Things to ask (the demo tasks)", ""]
        lines += [f"- **{t.title}**: {t.prompt}" for t in TASKS if t.workbook == DEMO_BOOK.name]
        lines += ["", "Reset the workbook at any time: `cp original/sales_q3.xlsx .`"]
        (ws / "PROMPTS.md").write_text("\n".join(lines) + "\n")

    status = subscription_status()
    print(f"workspace ready: {ws}")
    if status and status.get("loggedIn") and status.get("authMethod") == "claude.ai":
        print(f"Claude Code will use your claude.ai {status.get('subscriptionType', '')} "
              f"subscription ({status.get('email', '')}).")
    else:
        print("note: `claude auth status` does not show a claude.ai subscription login; run "
              "`claude auth login` and choose your claude.ai account to use your subscription.")
    if os.environ.get("ANTHROPIC_API_KEY"):
        print("note: ANTHROPIC_API_KEY is set in this shell; Claude Code may prefer it. Start with\n"
              "      `env -u ANTHROPIC_API_KEY claude` to use the subscription.")
    print(f"\nnext:\n  cd {ws}\n  claude")
    print("  -> on first start, accept the 'trust this folder' prompt, so the pre-approved\n"
          "     commands in .claude/settings.json take effect; then ask for edits in plain English.")
    if (ws / "PROMPTS.md").exists():
        print("  (sample requests are in PROMPTS.md)")
    print("\nnon-interactive (one request, then exit):\n"
          "  claude -p \"<your request>\" --settings .claude/settings.json")


if __name__ == "__main__":
    main()
