# Stability and patient-data protection — 7.9

## Approved scope

- Patient history now rejects malformed envelopes, unreadable encryption and invalid record containers. Reads never substitute an empty database for an unreadable file; direct compatibility saves also validate the current file before writing. Original bytes remain untouched on rejection.
- New patients are created without a record ID, including identical names. Updates and additional prescriptions for an existing patient require its explicitly selected ID. A missing/deleted ID fails rather than falling back to name matching. Previously merged historical records are not automatically split or migrated.
- Manual, Settings-triggered automatic and startup-background backups include validated encrypted patient history when present. Existing archives without a history entry leave live history unchanged. Restoration validates the history before replacing files and rolls back files already replaced when an ordinary later I/O operation fails. This is not a power-loss-safe multi-file journal.
- Startup displays a loading label without touching the derived SQLite cache. The worker rebuilds malformed/missing-table caches from the CSV, preserving the source. Loader failure releases navigation so Settings/import remain accessible.
- Medicine read/modify/write mutations are serialized by a shared per-path reentrant lock across database instances in this process. CSV writes use unique same-directory temporary files, fsync and replacement; SQLite staging names are unique. Repeated imports are blocked while pending. These are not cross-process or distributed locks.

## Preservation and limitations

No public API or cloud contract was removed. Existing DPAPI data, medicine source files and versioned executables are preserved. Backups remain encrypted for the originating Windows account; restoring patient history on a different account requires a separately designed migration, not copying encryption keys. Language rebuilds, export shutdown semantics and the remaining OpenFDA/cache issues from the audit are outside this release.

## Regression coverage

Isolated tests exercise corrupt history rejection with byte-for-byte preservation, same-name patient separation, explicit-ID updates, missing IDs, encrypted backup round trips, legacy archives, malformed archive history, rollback after a late restore failure, both automatic backup paths, invalid SQLite recovery, concurrent CSV merges across database instances, and duplicate import prevention. Existing API, storage, cloud and document tests are included in the release suite.

Verification commands (run from the project root with isolated data):

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\output\v79-release'
& E:\PDF\python.exe -m pytest -q --basetemp=.test-tmp-v79-release
```

The separate executable is `dist/RxPrescription-v7.9.exe`. Previous builds remain available.

Release verification: **210 tests passed in 75.77 seconds**. The packaged app
remained running during an isolated launch check, with no nonempty application
error log. This is not cross-Windows certification or a power-loss simulation.
