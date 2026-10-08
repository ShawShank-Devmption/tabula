"""Matched paired agent experiment. Freeze fixture/prompt/source manifest before execution.

No live agents run unless --driver claude is selected. --driver reference calibrates
both correct implementations (openpyxl and real Tabula TEL), not agent performance.
"""
from __future__ import annotations
import argparse
import json
import math
import os
import random
import shlex
import shutil
import signal
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

HERE=Path(__file__).resolve().parent
ROOT=HERE.parent
sys.path.insert(0,str(HERE))
from suite import TASKS, TASKS_BY_ID, SUITE_VERSION, SCORING_VERSION, freeze, sha256, solve, tel_solution
from run import subscription_env, require_subscription

COMMON='''You are working in an isolated evaluation folder containing book.xlsx.
Task: {task}

Use Python 3 and openpyxl (already installed) to inspect and edit the workbook. Run scripts with python3. Before editing, inspect the real headers, records, formulas, cross-sheet references and defined names. Map records by their labels rather than assuming row numbers. Plan the smallest edit; preserve unrelated content and formatting. Inspect for unsupported address-bearing features before structural edits. Verify the saved workbook and live formulas before claiming success. Stay inside this folder; do not access other task folders, grader, reference solutions or outside files. Save book.xlsx in place.

End your final response with exactly one machine-readable line:
EVAL_STATUS: SUCCESS
or EVAL_STATUS: REFUSAL
or EVAL_STATUS: UNVERIFIED
or EVAL_STATUS: FAILURE
Use SUCCESS only when you completed and verified the requested change. Use REFUSAL when you safely declined and left book.xlsx unchanged; explain why. Use UNVERIFIED if correctness could not be established. A normal process exit is not a success claim.'''
TABULA_NOTE='''

You additionally have ./tabula and TEL.md. Read the reference. Prefer Tabula for supported edits: inspect, write TEL, check the edit, inspect diagnostics, apply --in-place only after the check succeeds, and verify. Python/openpyxl remains available for inspection or safe fallback. Report which method you actually used.'''
COMMON_ALLOWED=['Bash(python3 *)','Bash(python *)','Bash(ls *)','Bash(cat *)','Bash(head *)','Bash(tail *)','Bash(grep *)','Bash(wc *)','Read','Write','Edit','Glob','Grep']
TOOLS='Bash,Read,Write,Edit,Glob,Grep'


def prompt(task,arm):return COMMON.format(task=task.prompt)+(TABULA_NOTE if arm=='tabula' else '')


def parse_args(argv=None):
    ap=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--driver',choices=('reference','claude'),default='reference')
    ap.add_argument('--tasks',default=','.join(t.id for t in TASKS))
    ap.add_argument('--arms',default='default,tabula')
    ap.add_argument('--trials',type=int,default=3)
    ap.add_argument('--model',default='claude-sonnet-5-5')
    ap.add_argument('--budget',type=float,default=2.0)
    ap.add_argument('--timeout',type=int,default=600)
    ap.add_argument('--parallel',type=int,default=1)
    ap.add_argument('--seed',type=int,default=20261008)
    ap.add_argument('--out',type=Path)
    args=ap.parse_args(argv)
    args.tasks=args.tasks.split(',');args.arms=args.arms.split(',')
    for field,allowed in (('tasks',TASKS_BY_ID),('arms',('default','tabula'))):
        value=getattr(args,field)
        if not value or len(set(value))!=len(value) or any(v not in allowed for v in value):ap.error(f'invalid or duplicate {field}')
    if args.trials<1 or args.timeout<1 or args.parallel<1 or not math.isfinite(args.budget) or args.budget<=0:ap.error('trials, timeout, parallel and budget must be positive and finite')
    if not args.model.strip():ap.error('model must be nonempty')
    args.out=(args.out or HERE/'runs'/f"{datetime.now():%Y%m%d-%H%M%S}-suite-{args.driver}").resolve()
    if args.out.exists():ap.error(f'output already exists (no overwrite): {args.out}')
    return args


def bounded_process(cmd,cwd,stdout,stderr,timeout,env):
    """Kill the whole spawned process group on timeout, not only the CLI parent."""
    start=time.monotonic()
    with Path(stdout).open('w') as out,Path(stderr).open('w') as err:
        proc=subprocess.Popen(cmd,cwd=cwd,stdout=out,stderr=err,stdin=subprocess.DEVNULL,env=env,start_new_session=True)
        try:code=proc.wait(timeout=timeout);timed_out=False
        except subprocess.TimeoutExpired:
            os.killpg(proc.pid,signal.SIGKILL);proc.wait();code=None;timed_out=True
    return dict(returncode=code,timed_out=timed_out,duration_s=round(time.monotonic()-start,3))


def transcript_metadata(path):
    tool_calls=0;tabula_used=False;result=None;resolved=set();calls=[]
    for line in Path(path).read_text().splitlines():
        try:event=json.loads(line)
        except json.JSONDecodeError:continue
        if event.get('type')=='assistant':
            model=event.get('message',{}).get('model')
            if model:resolved.add(model)
            for item in event.get('message',{}).get('content',[]):
                if isinstance(item,dict) and item.get('type')=='tool_use':
                    tool_calls+=1
                    cmd=item.get('input',{}).get('command','')
                    if item.get('name')=='Bash' and ('./tabula ' in cmd or '-m tabula ' in cmd):tabula_used=True;calls.append(cmd)
        elif event.get('type')=='result':result=event
    meta=dict(tool_calls=tool_calls,tabula_used=tabula_used,tabula_commands=calls,model_ids=sorted(resolved))
    if result:
        meta.update(is_error=bool(result.get('is_error')),final_message=result.get('result') or '',
            cost_usd=result.get('total_cost_usd'),num_turns=result.get('num_turns'),usage=result.get('usage'),model_usage=result.get('modelUsage'),
            permission_denials=result.get('permission_denials') or [],result_subtype=result.get('subtype'))
    else:meta.update(is_error=True,final_message='',runner_error='No final result event in transcript')
    return meta


def source_hashes():
    paths=list((ROOT/'tabula').rglob('*.py'))+[ROOT/'docs/TEL.md',HERE/'suite.py',HERE/'suite_run.py',HERE/'suite_evaluate.py',HERE/'run.py']
    return {str(p.relative_to(ROOT)):sha256(p) for p in sorted(paths)}


def one(job,args):
    task=TASKS_BY_ID[job['task']];rundir=args.out/job['directory']
    rundir.mkdir()
    shutil.copy(args.out/'fixtures'/f'{task.id}.xlsx',rundir/'book.xlsx')
    (rundir/'prompt.txt').write_text(prompt(task,job['arm']))
    if job['arm']=='tabula':
        shutil.copy(ROOT/'docs/TEL.md',rundir/'TEL.md')
        # POSIX quoting avoids paths with spaces becoming shell commands.
        (rundir/'tabula').write_text(f'#!/bin/sh\nPYTHONPATH={shlex.quote(str(ROOT))} exec {shlex.quote(sys.executable)} -m tabula "$@"\n')
        (rundir/'tabula').chmod(0o755)
    meta=dict(job,driver=args.driver,model=args.model if args.driver=='claude' else None,
              input_sha256=sha256(rundir/'book.xlsx'),started_utc=datetime.now(timezone.utc).isoformat())
    try:
        if args.driver=='reference':
            if task.operation.startswith('refuse'):
                cmd=[sys.executable,'-c','print("Certified relocation unavailable; workbook untouched.\\nEVAL_STATUS: REFUSAL")']
                used=False
            elif job['arm']=='tabula':
                (rundir/'edit.tel').write_text(tel_solution(task))
                cmd=[str(rundir/'tabula'),'apply','book.xlsx','edit.tel','--in-place'];used=True
            else:
                cmd=[sys.executable,str(HERE/'suite_reference.py'),task.id];used=False
            status=bounded_process(cmd,rundir,rundir/'transcript.txt',rundir/'stderr.txt',args.timeout,dict(os.environ,PYTHONPATH=str(ROOT)))
            meta.update(status,is_error=status['returncode']!=0,tabula_used=used,tool_calls=1,
                final_message=('EVAL_STATUS: REFUSAL' if task.operation.startswith('refuse') else 'EVAL_STATUS: SUCCESS') if status['returncode']==0 else 'EVAL_STATUS: FAILURE',cost_usd=None,usage=None)
        else:
            allowed=COMMON_ALLOWED+(['Bash(./tabula *)'] if job['arm']=='tabula' else [])
            cmd=['claude','-p',prompt(task,job['arm']),'--safe-mode','--strict-mcp-config','--no-session-persistence',
                '--model',args.model,'--output-format','stream-json','--verbose','--max-budget-usd',str(args.budget),
                '--permission-mode','dontAsk','--permission-prompts','none','--tools',TOOLS,'--allowedTools',*allowed]
            meta.update(bounded_process(cmd,rundir,rundir/'transcript.jsonl',rundir/'stderr.txt',args.timeout,subscription_env()))
            meta.update(transcript_metadata(rundir/'transcript.jsonl'))
        meta['output_sha256']=sha256(rundir/'book.xlsx')
    except Exception as exc:meta.update(returncode=None,is_error=True,runner_error=f'{type(exc).__name__}: {exc}',final_message='')
    (rundir/'meta.json').write_text(json.dumps(meta,indent=2)+'\n')
    print(f"Finished {job['directory']} ({meta.get('duration_s','?')}s)",flush=True)
    return meta


def main(argv=None):
    args=parse_args(argv)
    checker=HERE/'.venv/bin/python'
    if not checker.exists():raise SystemExit('Missing demo/.venv checker environment; install demo/requirements.txt first')
    check=subprocess.run([str(checker),'-c','import pycel,openpyxl; print(openpyxl.__version__)'],capture_output=True,text=True,timeout=15)
    if check.returncode:raise SystemExit('Independent checker dependencies unavailable: '+check.stderr)
    auth=None;version='scripted-reference'
    if args.driver=='claude':
        auth=require_subscription()
        version=subprocess.run(['claude','--version'],capture_output=True,text=True,timeout=15,env=subscription_env(),check=True).stdout.strip()
    args.out.mkdir(parents=True,exist_ok=False)
    frozen=freeze(args.out/'fixtures')
    rng=random.Random(args.seed)
    pairs=[(task,trial) for task in args.tasks for trial in range(1,args.trials+1)]
    rng.shuffle(pairs);jobs=[]
    for task,trial in pairs:
        arms=list(args.arms);rng.shuffle(arms)
        for arm in arms:jobs.append(dict(task=task,arm=arm,trial=trial,directory=f'{task}__{arm}__{trial}',launch_order=len(jobs)))
    manifest=dict(suite_version=SUITE_VERSION,scoring_version=SCORING_VERSION,checker_sha256=sha256(HERE/'suite_evaluate.py'),
        frozen_utc=datetime.now(timezone.utc).isoformat(),driver=args.driver,model=args.model if args.driver=='claude' else None,runner_version=version,
        budget_usd=args.budget,timeout_s=args.timeout,parallel=args.parallel,seed=args.seed,arms=args.arms,trials=args.trials,
        tasks=[t for t in frozen['tasks'] if t['id'] in args.tasks],jobs=jobs,source_hashes=source_hashes(),
        python_version=sys.version,checker_openpyxl=check.stdout.strip(),auth={'authMethod':auth.get('authMethod'),'subscriptionType':auth.get('subscriptionType')} if auth else None,
        prompts={f'{t.id}__{a}':prompt(t,a) for t in TASKS if t.id in args.tasks for a in args.arms},
        allowed_tools={a:COMMON_ALLOWED+(['Bash(./tabula *)'] if a=='tabula' else []) for a in args.arms})
    (args.out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
    print(f'Frozen manifest; {len(jobs)} runs -> {args.out}',flush=True)
    with ThreadPoolExecutor(max_workers=args.parallel) as pool:
        list(pool.map(lambda j:one(j,args),jobs))
    # Refuse to grade a moving implementation as if it were the frozen experiment.
    if source_hashes()!=manifest['source_hashes']:raise SystemExit('Source files changed during execution; run invalid, preserved for inspection')
    subprocess.run([str(checker),str(HERE/'suite_evaluate.py'),str(args.out)],check=True)

if __name__=='__main__':main()
