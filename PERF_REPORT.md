# Performance implementation — version 7.3

## Follow-up release 7.4

Favorite and saved-template batches now pause beyond the viewport plus a 160-pixel buffer, resume as scrolling approaches them, and automatically request subsequent batches when their footer enters that buffer. Existing Load more controls remain available. Generation guards discard superseded checks, and hidden pages pause their work. This is incremental construction, not full widget recycling: already visited favorite cards remain cached.

Window-size changes defer glass rasterization until 200 ms after resizing settles, retaining the existing image during the gesture. Replacement templates reuse available medication rows, invalidate pending autocomplete results, overwrite every medication field and destroy surplus rows. New rows are constructed only when the replacement needs more than the existing prescription contains.

Added regression coverage for viewport buffering, resize delay, medication-row identity reuse and removal of surplus rows. No new storage format, public API or export-layout change. Earlier versioned executables are retained. The 7.3 timing figures below are historical; no universal 7.4 speedup or startup-budget claim is made.

Implemented the approved P01–P11 audit changes while preserving the desktop UI, public interfaces, cloud contract, encrypted storage formats and search result ordering. Earlier versioned executables remain available.

## Changes

- Glass images are clipped to the visible region and rasterized on a dedicated worker. Tk image creation stays on the UI thread. A bounded strong cache avoids redundant rendering; obsolete queued work is cancelled.
- Navigation updates only the previous and next pages, preserves browsing state, and avoids remapping the current page.
- Template insertion yields between rows and performs numbering/final refresh once. Favorites and saved templates reuse cards; hidden-page rendering pauses until revisited. Template action callbacks use the latest record even when timestamp-only changes reuse a card.
- Autocomplete debounce is 35 ms. A derived SQLite search index improves combined-name lookup without changing normalization, prefix-first ordering or limits.
- Prescription saves capture current data and run on a serialized worker. Repeated saves are blocked while pending, and newer edits are not marked saved by an older operation. Closing is blocked while an explicit save remains pending.
- Patient history uses an invalidated in-memory cache, copies only requested records where possible, and serializes in-process transactions. Patient list/history rendering yields in batches.
- Settings and patient history writes use unique same-directory temporary files, flush/fsync and atomic replacement. Unchanged settings saves skip redundant encryption only when the backing file is unchanged. DPAPI and existing envelope formats are unchanged.
- Database loading is deferred, with dependent navigation disabled until ready. PDF imports are lazy and remain included in packaging. Worker completion handoffs, including openFDA, use the UI queue rather than Tk calls from threads.

## Verification

- Final isolated regression suite: **200 passed in 110.68 seconds**.
- New performance tests cover cache invalidation, defensive copies, concurrent history writes, atomic-write failure, search equivalence, clipped rendering, card reuse, UI-thread handoffs, asynchronous snapshots, repeated-click protection and template ordering.
- Synthetic tests generated English/Arabic A4/A5 PDF and both Word variants. Existing cloud, layout and public-interface regressions passed. No real patient upload or live network reference lookup was needed for this performance change.
- Built `dist/RxPrescription-v7.3.exe`. An isolated process-launch smoke test remained running for 12 seconds without an application error log. Archive inspection confirmed PDF, ReportLab, Word, history and database dependencies. This is not a complete interactive frozen-export test or a cross-Windows certification.
- Measurements are in `output/perf-audit-v72/` and `output/perf-audit-v73/`; methodology and initial findings remain in `PERF_BASELINE.md` and `PERF_AUDIT.md`.

## Timing interpretation and remaining limits

The initial implementation run reduced unchanged favorite refresh from approximately 838 ms to 25 ms and template refresh from 710 ms to 22 ms. Autocomplete appeared in approximately 62–86 ms in that run, versus 203–238 ms initially. Prescription-save handler time fell from approximately 1,751 ms to 1.3 ms because the expensive operation moved off the UI thread; total completion is separately recorded and remains much longer.

The final-source run measured favorite/template refresh at 32/30 ms, prescription-save handler at 19 ms, and checked autocomplete samples at 66/82 ms (English/Arabic). Ten template rows completed in 4.57 seconds and the save/UI completion in 9.18 seconds. The first update was 1.92 seconds excluding imports, not a cold usable-start guarantee. Final-source UI measurements are in `output/perf-audit-v73/ui.json`; compare raw samples rather than treating asynchronous return times as completed rendering. Baseline debug instrumentation and initial overlapping work mean these figures are directional, not controlled universal speedup claims.

**Not all requested budgets are met.** Initial card creation, first startup/layout and scripted resizing still take seconds on the large fixture; a single widget-heavy callback can exceed 50 ms. Template insertion is cooperative, not instantaneous, and adding one row can still exceed 100 ms. Patient basic-save/delete handlers and large old-card destruction remain synchronous. The glass-cache limit excludes images held by currently visible widgets, and tracemalloc excludes native Pillow/Tk allocations. Whole-file history encryption still costs time on the worker. Cross-process history transactions are not fully locked; detected external replacements are rejected rather than silently overwritten.

No OS-reboot cold-start, native-memory profile, slow-disk simulation, physical printing or multi-Windows-version test was performed. One-file production packaging is retained because the earlier exploratory comparison did not establish an onedir advantage.

## Generated-artifact cleanup — 2026-10-02

Removed temporary test directories, bytecode, PyInstaller intermediates, synthetic benchmark databases and generated sample exports after verification, reclaiming approximately 0.60 GiB. Raw benchmark JSON/import traces, reports, source, design assets, backend checkout and all versioned executables remain. Re-running the benchmark requires its `seed` mode first. Application version remains 7.3 because this cleanup changes no application code or behavior.
