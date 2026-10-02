# Stability update 8.0

Scope: selected remaining audit items 1, 2, 3, 4 and 8. Existing APIs,
patient records and earlier executables are preserved.

## Behavior

- Changing interface language captures patient, prescriber and medication
  inputs in memory before rebuilding the forms. The selected patient identity,
  saved-workflow signature and expanded medication are restored. This is not
  session recovery: reopening the application still starts a fresh prescription.
- Loaded and restored settings are checked against default section/field types.
  Malformed sections receive safe defaults; unrelated valid fields, including
  the cloud key, are retained. Null template medication collections are safe.
- Normal window closure is blocked while an export is active. Finish the export,
  or choose Cancel in its upload-failure dialog, before closing. Settings cannot
  rebuild the interface during an export or patient-history save. This does not
  provide recovery from forced process termination or power loss.
- OpenFDA callbacks carry a request generation. Superseded responses and callbacks
  after closing cannot update the result panel. Worker code never touches Tk.
- Favorites and saved templates recycle heavy off-screen contents when a pool
  exceeds 96 live cards. Lightweight geometry spacers preserve ordering and
  scrolling; entering the viewport rebuilds contents/actions from their records.
  This bounds expensive contents, not total record or spacer count.

## Verification

Use an isolated application-data directory for all tests and smoke runs:

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\output\v80-test'
& E:\PDF\python.exe -m pytest --basetemp=.test-tmp-v80-final -q --tb=short
```

Six new regression checks cover language/input preservation and a fresh next
launch; malformed settings and key retention; export-close protection; latest
OpenFDA callback and closure protection; favorite-card recycling and restored
actions; and template-card recycling and restored expandable medicine contents.

Package separately using `build_exe.ps1`; the version-derived output is
`dist/RxPrescription-v8.0.exe`. No live patient data or authenticated network
uploads are required for this update's verification.

Release result: **216 tests passed in 111.67 seconds**. The actual Arabic
language-change check also passed separately. The packaged executable remained
running for the isolated 10-second launch check with no nonempty application
error log. Previous versioned executables remain present. These checks are not
cross-Windows certification or forced-termination recovery guarantees.
