# Project status

Updated: 2026-10-03. Application version: **8.0**.
Repository: https://github.com/ahmedinternist/rx-prescription, branch `main`.
Implementation commit: `df369a9011df80ae316d826a641d8f2a6ed5e624`.
This status update is documentation-only; it does not change the executable version.

## Completed

### Latest release: 8.0

Implemented the selected remaining audit items **1, 2, 3, 4 and 8**:

1. Preserve current patient/prescriber inputs, medications, selected patient ID,
   saved-workflow signature and expanded medication across interface-language
   changes. Capture is session-only: unfinished prescriptions are not saved or
   restored on the next launch. Cancel old scheduled form work and reject old
   card-render callbacks when rebuilding the interface.
2. Validate loaded/restored settings sections and field types. Repair malformed
   values using defaults while retaining unrelated valid values, including the
   cloud API key. Handle null template medication lists safely.
3. Block normal window closure during active exports. Also block Settings saves
   that would rebuild the interface during an export or patient-history save.
   Finish the export, or cancel in its upload-failure dialog, before closing.
4. Reject superseded OpenFDA callbacks using request generations; ignore
   callbacks after shutdown. Worker-to-Tk delivery remains queue-based.
5. Recycle heavy off-screen favorite/template contents above 96 live cards per
   pool. Preserve lightweight geometry spacers and recreate actions/details
   when cards are needed again.

Built `E:\Prescription App\dist\RxPrescription-v8.0.exe` separately and retained
previous versioned builds. Source, tests, documentation and accumulated earlier
changes were committed and pushed to `origin/main`. Executables and generated
test/output directories are ignored by Git and were not uploaded.

### Preceding data-protection release: 7.9

- Fail closed on unreadable/corrupt patient history without overwriting it.
- Keep identical patient names separate; updates require the selected record ID.
- Include encrypted patient history in manual and automatic backups; validate
  before restore and roll back ordinary later I/O failures.
- Rebuild invalid derived medicine caches without modifying the source CSV.
- Serialize in-process medicine database mutations and reject overlapping imports.

The same pushed commit also includes accumulated workflow, print-layout,
cloud/configuration, visual, performance and cleanup work. The historical reports
describe their original versions, not fresh measurements of 8.0.

## Files changed

The complete file inventory for implementation commit `df369a9` is **47 files**:

| Group | Files |
|---|---|
| Modified application/storage/export | `main.py`, `config.py`, `drug_db.py`, `patient_history.py`, `cloud_rx.py`, `qr_utils.py`, `pdf_generator.py`, `workflow_features.py`, `i18n.py` |
| Modified build/test configuration | `.gitignore`, `build_exe.ps1`, `requirements.txt`, `pytest.ini` |
| Modified tests | `test_core.py`, `test_cloud_rx.py`, `test_workflow_features.py` |
| Modified documentation | `README.md`, `Cloud_QR_Verification.md`, `Font_Groups_Verification.md`, `Windows_Compatibility_Report.md` |
| Added application/diagnostics | `print_layout.py`, `perf_probe.py`, `perf_baseline.py`, `perf_package_probe.ps1` |
| Added tests | `test_api_compatibility.py`, `test_performance.py`, `test_print_layout.py` |
| Added reports | `PERF_AUDIT.md`, `PERF_BASELINE.md`, `PERF_REPORT.md`, `Stability_7.9_Verification.md`, `Stability_8.0_Verification.md`, `Visual_Styles_Verification.md` |
| Deleted retired helpers/samples | `_check_url.py`, `_gen_scan_help.py`, `_gen_test.py`, `scan_help.pdf`, `test_qr.png`, `viewer.html`, `viewer_test.html`, `test_headless.py`, `test_new_features.py` |
| Deleted retired reference/compatibility implementation | `gemini_drug.py`, `backend-compat/README.md`, `backend-compat/rx-normalizer.ts`, `backend-compat/test-normalizer.mjs`, `backend-compat/viewer.patch` |

The 8.0-specific edits are in `main.py`, `config.py`, `i18n.py`, `test_core.py`,
`README.md` and `Stability_8.0_Verification.md`. This follow-up adds
`PROJECT_STATUS.md`; it does not modify application code.

## Tests and checks run

All application imports/test launches used isolated `RX_APP_DATA_DIR` directories,
not the user's real patient records or encrypted configuration.

| Check | Recorded result |
|---|---|
| 7.9 full regression suite | 210 passed in 75.77 seconds |
| 8.0 initial existing suite | 210 passed in 73.07 seconds |
| 8.0 targeted new regressions | 5 passed; subsequently expanded to six new tests |
| 8.0 full suite before final render-callback guard | 216 passed in 121.15 seconds |
| New tests plus lazy-page/reload check | 7 passed in 31.78 seconds |
| Actual Arabic language-change test, using `I.set_lang("ar")` | 1 passed in 6.54 seconds |
| Final 8.0 full regression suite | **216 passed in 111.67 seconds** |
| PyInstaller production build | `build_exe.ps1` completed successfully; v8.0 executable produced |
| Packaged executable isolated launch | Running after 10 seconds; no nonempty application error log; only the test-launched process tree was terminated |
| Git validation before implementation commit | `git diff --check` and staged diff check passed; common credential-pattern scan found no matches in changed files |
| Push verification | Local HEAD and `origin/main` both matched `df369a9`; working tree clean |

Final full-suite command:

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\output\v80-test'
& E:\PDF\python.exe -m pytest --basetemp=.test-tmp-v80-release -q --tb=short
```

Additional actual-language check:

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\output\v80-test'
& E:\PDF\python.exe -m pytest test_core.py -k language_change_preserves --basetemp=.test-tmp-v80-real-language -q --tb=short
```

These results were obtained during implementation, not rerun for this
documentation-only update. The status commit must pass `git diff --cached --check`
before pushing.

## Unresolved issues and verification limits

- Remaining audit items outside the selected scope are not claimed fixed.
  Re-audit the current source before assigning new priorities; the old performance
  reports and line numbers refer to earlier snapshots.
- No cross-Windows certification: this release has not been exercised on Windows
  10/Server/ARM test machines or a multi-version VM matrix.
- No physical A4/A5 printing or interactive full export acceptance test was
  performed for 8.0. Regression tests cover the existing document/cloud paths;
  the launch smoke check is not a full UI walkthrough.
- No fresh authenticated live cloud upload/viewer scan or live OpenFDA request
  was performed for this stability release. Network tests use mocks.
- Export-close protection covers normal application closure, not task killing,
  power loss or recovery of unfinished exports.
- Database serialization is in-process only. Backup restoration has ordinary
  I/O rollback, not a power-loss-safe multi-file transaction journal.
- DPAPI history/backups remain bound to the originating Windows account.
  Cross-account migration and splitting previously merged historical patient
  records remain separate work.
- The card limit bounds heavy widget contents, not total records/spacer count or
  all native-memory allocations. A fresh 8.0 performance/native-memory baseline
  is still needed; earlier responsiveness measurements are not current guarantees.

## Exact next step

Perform a **manual, fictitious-data acceptance pass of the packaged 8.0 app**
before another feature change. Start on the current PC using a fresh isolated
profile; do not copy the real configuration or patient history:

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\output\v80-manual-acceptance'
Start-Process -FilePath 'E:\Prescription App\dist\RxPrescription-v8.0.exe'
```

1. Enter a fictitious Arabic patient and two medicines; change language and
   verify input, sex selection, ordering and expansion survive. Close/reopen and
   verify the unfinished prescription is absent.
2. Browse enough favorites/saved templates to exercise off-screen recycling;
   scroll away/back and check selection, expansion and Add/Edit/Delete actions.
3. Verify rapid successive OpenFDA searches display only the latest medicine.
4. With a test-only cloud key supplied manually in Settings, export fictitious
   A4/A5 PDF and both Word variants. Try closing during export, and verify the
   failure dialog's Retry/Export without QR/Cancel choices using a controlled
   network failure. Never place the key or prescription bodies in the report.
5. Repeat the acceptance pass on a Windows 10 x64 test PC with its own isolated
   profile. Record OS/build/DPI, pass/fail results and any sanitized error traces
   in a new release-acceptance report. Only then re-triage unselected audit items.

The next step is verification, not authorization to change storage, authentication,
cloud contracts or resume unfinished prescriptions.

## Shared Codex/Claude handoff

The remote handoff infrastructure from `c2ceb04`, `80e9fbb` and `7df4216` was
integrated without altering `AGENTS.md` or `CLAUDE.md`. Both histories are
preserved; the add/add conflict in this file was resolved by combining the
detailed release report with the shared handoff context.

The application is a bilingual English/Arabic Python Windows desktop UI for
prescription entry/history, PDF/Word export, database search, reusable templates
and the deployed `rx-v2` Next.js/Redis short-link viewer. Local sensitive data is
DPAPI-protected. The existing cloud payload contract remains unchanged; new links
use the configured 60-day retention policy from the 5.7 backend deployment.

No unfinished coding task is assumed. The next task is the acceptance pass above.
Before future changes, read `README.md`, this file, `AGENTS.md` and `CLAUDE.md`,
then inspect Git status/history and relevant code. Record the task objective,
constraints and acceptance criteria; update completed work, changed files,
verification, unresolved risks and the exact next step before committing/pushing.
Never weaken DPAPI, patient privacy, cloud secrets/contracts, Arabic/RTL,
PDF/Word or database/import compatibility without explicit task authorization.

This handoff changed documentation only. Conflict-marker and staged-diff checks
are required before completing the merge and pushing; no runtime code from the
8.0 implementation was changed by integrating these remote documents.
