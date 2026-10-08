# Safety and agent evaluation implementation ledger

Authorized scope: the twelve-task list approved in conversation on 2026-10-08.
Baseline: e9080e3; `python3 -m pytest tests/ -q`: 162 passed (1.04s).

## Execution constraints and decisions
- Git metadata is read-only in this sandbox; branch creation failed. Work in the existing checkout, leave changes uncommitted, record the baseline and final source hashes for reproducibility.
- Preserve current compiler/engine/XLSX architecture; no new language or VM.
- Unsupported address-bearing workbook features must cause safe refusal instead of guessing relocation.
- Conservatively invalidate unsupported formula dependencies on edits. Unverified expectations remain explicitly unverified, never PASS.
- Source snapshot and cooperating-writer locking plus commit-time revalidation protect apply. Document limits with external writers that ignore locks; do not claim portable filesystem compare-and-swap.
- Independent evaluation uses pycel and Python expectations, never Tabula to judge Tabula.
- Comparison arms share Python/openpyxl and workflow guidance; Tabula arm additionally receives Tabula and its reference. Freeze scenarios and manifest before measured runs.
- No real Excel-derived caches will be fabricated. If Excel calculation is unavailable, provide a corpus validator and report the missing external evidence.
- Live runner was not signed in during implementation; measured runs are deferred to Phase 3.

## Tasks
Status as of the checkpoint commit (full suite: `python3 -m pytest tests/ -q` -> 219 passed).

- [x] 1. Reproducible baseline recorded; suite runs freeze a manifest with source hashes.
- [x] 2. Conservative uncertainty propagation (`tests/test_dependency_safety.py`).
- [x] 3. SUMIF effective dependency ranges and recompute equivalence (randomised
      incremental-vs-full test).
- [x] 4. Structural/rename reference safety: workbook-wide refusal (`tests/test_boundary_safety.py`).
- [x] 5. Grid-boundary validation and structural output verification.
- [x] 6. Consistent file snapshot and concurrency guards (residual limit documented below).
- [x] 7. Workbook date epoch (1900/1904).
- [~] 8. Regression suite done. External differential corpus still open: the two files in
      `tests/fixtures/excel/` were written by openpyxl, carry no Excel-calculated values,
      and are not yet used by a test. They must be opened and saved in Excel first.
- [x] 9. Independent evaluator preservation/mutation checks (`tests/test_demo_suite.py`).
- [x] 10. Fifteen evaluation tasks across five workbooks (`demo/suite.py`).
- [~] 11. Matched runner implemented (`demo/suite_run.py`). Reference calibration
      (`--driver reference --trials 1`): 26 successful completions + 4 safe refusals of 30,
      as designed. Measured live runs are not yet executed.
- [~] 12. Documentation synced with the implementation; comparison report and independent
      final review pending the live runs.

## Ownership
- Engine/dependency analysis, documentation/integration: primary agent.
- XLSX boundary/compiler apply and structural safety in simulate: boundary implementer.
- Evaluation tasks, runner, independent checker and demo tests: evaluation implementer.
- Shared interface: engine propagates unverified values before simulate builds its report; boundary implementer includes unparsed cells in unverified reporting.
