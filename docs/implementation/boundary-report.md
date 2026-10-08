# XLSX boundary safety implementation

Implemented in `tabula/xlsx.py`, `tabula/tel/compiler.py`, `tabula/tel/simulate.py`; regression coverage in `tests/test_boundary_safety.py`. Existing public call signatures remain usable; `write_back` adds an optional destination snapshot parameter and `Loaded` adds an optional source-parts field.

## Reference safety

Structural edits and sheet renames now return `E-STRUCT` when any workbook sheet contains unsupported address-bearing features. Counts include charts, pivot tables, chart sheets, external links, merged ranges, tables, conditional formatting, validations, hyperlinks, print areas/titles, and autofilters. This deliberately refuses even features on other sheets and even when a simple textual reference search finds no direct link: indirect name references and address strings cannot be ruled out safely.

Unparsed formulas and defined names block both structural edits and renames. Unsupported parsed formulas and names also block them, including `INDIRECT("Lists!A1")`, whose string address cannot be relocated through the AST. This trades some usable static-reference workbooks for conservative correctness. Refusing a rename stops subsequent operations because the resolver has already bound them against the proposed new sheet name; it produces a diagnostic instead of crashing or applying an incomplete edit.

Dirty unparsed formulas appear in the unverified report. Formula relocation counts still count parsed formulas only. Simulation passes the engine's new `changed` flag when an edit or structure operation occurred, including edits with no ordinary formula seeds.

## Grid and backend agreement

Loading records sparse stored cell positions, including style-only cells, plus row dimension positions and column dimension spans. Simulation moves this metadata alongside model values and rejects inserts that would push any stored position or formatting dimension beyond `XFD` or row `1048576`. Prior deletes and new writes are taken into account.

After structural replay the backend verifies every model cell, including untouched values that merely moved, and refuses any unexpected nonempty output cell. Existing touched-cell verification remains active. Verification converts reloaded dates using the output workbook's epoch. The reported synced count continues to count cells actually rewritten, not all cells examined for verification.

## Snapshot and concurrency

Loading reads one immutable byte snapshot. Formula loading, cached-value loading, source SHA-256, and fidelity source-part counts all derive from those same bytes. A concurrent save between the two openpyxl reads therefore cannot mix workbook versions.

`apply` acquires stable POSIX `fcntl.flock` sidecar locks for both source and output before loading/simulating. Paths are canonicalized and locks acquired in sorted order. Sidecars remain in place: deleting them would allow another writer to lock a new inode while the original inode remains locked. Platforms without `fcntl` fail clearly with `E-LOCKED` rather than proceeding without coordination.

The output's existence/hash is captured before simulation. Immediately before `os.replace`, after save/fidelity/verification, the backend checks Excel owner files for source and output, rehashes the source against the loaded snapshot, and checks output existence/hash against its initial snapshot. Thus an existing output, newly created output, or source saved externally during check/save/verify is preserved with `E-CONFLICT`; a newly appearing owner file produces `E-LOCKED`. Temporary output files are cleaned on refusal.

**Residual limit:** sidecar locks serialize cooperating `apply` writers only. Excel and other tools may ignore them. The final checks and `os.replace` are separate filesystem operations; an external writer can still save or open the workbook in that short interval. This is not portable filesystem compare-and-swap, nor a claim that the external-writer race has been eliminated. Low-level `write_back` performs final revalidation but callers using it directly must provide their own cooperating-writer lock scope.

Dates are converted with `to_excel(value, epoch=xl.epoch)`, covering both literal values and loaded caches and preserving the 1900/1904 date system distinction.

## Verification evidence

- Initial regression run before production fixes: 21 failed, 1 passed. Failures demonstrated unsafe DV rename, stale chart-reference allowance, unparsed rename allowance, both grid edges for values/styles/dimensions, mixed-version reads, lost concurrent updates, missed commit-time owner files, and incorrect 1904 dates. The already passing test characterized existing style/dimension movement.
- A second regression cycle exposed synced-count inflation and parsed `INDIRECT` bypasses before those fixes; a separate regression exposed the crash after a refused rename.
- Temporarily disabling expanded structural output verification produced the expected failures for corruption of an untouched moved value and introduction of an unexpected cell; restoring it passed both.
- Targeted command: `python3 -m pytest tests/test_boundary_safety.py tests/test_tel.py -q`: **65 passed**.
- First full-suite run while evaluation work was in progress: **202 passed, 5 failed**. All failures were in `tests/test_demo_suite.py`: missing `suite`/`suite_evaluate` modules in four tests and the unrelated salesperson-corruption checker in `test_old_tax_demo_unrelated_salesperson_corruption_rejected`. These files are owned by the evaluation implementer; no boundary failures occurred.
- Final full-suite result is recorded below after integration verification.
