# Performance audit — approval required before fixes

Version 7.2, 2026-10-01. See [PERF_BASELINE.md](PERF_BASELINE.md) for stack, fixture, measurements and limitations. This audit adds diagnostics only. No functional fix, schema migration, encryption change, search change or new release was implemented.

Sorted approximately by expected impact relative to effort, with risk and measurement confidence considered. Costs below are observations, not guaranteed savings.

| ID | Area | File:Line | Problem | Measured cost | Impact | Effort | Risk to correctness/data |
|---|---|---|---|---|---|---|---|
| P01 | Glass rendering | `main.py:335`, `main.py:215` | Rasterizes full panel dimensions on UI thread; tall scroll content and repeated region/size changes generate large images. Existing weak cache does not retain images across unmap/revisit. | Individual paint up to 2.16 s; ten scripted resize steps 16.1 s | High | Medium | Visual fidelity/scroll alignment; **no storage/search** |
| P02 | Navigation | `main.py:2325` | Unmaps all pages and remaps target on every call, including same-page refresh; triggers glass release/repainting and repeated layout. | Warm page switches still roughly 1–2.4 s including paints | High | Small | Must preserve scroll restoration, page state and pending editor guards; **no storage/search** |
| P03 | Template application | `main.py:4938`, `main.py:4996` | Adds rows one at a time, repeatedly collapsing/renumbering rows and triggering traced updates instead of one batch. | Ten drugs 1.98 s; single row 147 ms | High | Medium | Medication order/values and preview must be identical; **no storage-format change** |
| P04 | Autocomplete scheduling | `main.py:1663` | 180 ms debounce alone exceeds the requested 100 ms key-to-results target despite very fast indexed queries. | 203–238 ms visible; actual indexed queries <6 ms | Medium | Small | **SEARCH TIMING** only; preserve exact matching/order, cancellation and stale-result guards |
| P05 | Patient persistence | `patient_history.py:27`, `patient_history.py:40`, `main.py:8064` | Prescription save decrypts/re-encrypts full history synchronously in UI handler; repeated reads also decrypt entire file. | UI save 1.75 s; history load 0.90 s; service save 1.00 s median | High | Medium–Large | **STORAGE/ENCRYPTION: approval required.** Keep DPAPI and exact format; serialize writes, preserve external-change detection and crash safety |
| P06 | Favorites/templates rendering | `main.py:5116`, `main.py:3772` | Favorites update many existing widget properties; template browsing reconstructs visible cards rather than updating only changed cards. Existing caps help but do not prevent costly batches. | Refreshes 838 / 710 ms | High | Medium | Card identity, expansion, selection and scroll preservation; **no storage/search** |
| P07 | Patient list/history rendering | `main.py:7693`, `main.py:7882` | All patient rows and all prescription header cards are built in one UI turn; collapsed bodies do not cap header creation. | 2,200-row list 115 ms; 300 history cards 1.40 s including load | Medium | Medium | Incremental UI batches must not omit records or reorder results; **no storage change** |
| P08 | openFDA handoff | `main.py:8342` | Worker calls `self.after` directly. `after` itself is a Tk call and can block/error during shutdown; use the existing queue handoff instead. | Not reproduced/timed; verified call path | Medium | Small | Thread-safety/shutdown fix; **no API removal, storage or search changes** |
| P09 | Startup | `main.py` App initialization; `drug_db.py:316` | A stale/missing cache rebuild occurs before usable UI. ReportLab/PDF import also contributes to startup, while Word is already lazy. | Cache rebuild 1.65 s; PDF import 162 ms | Medium | Medium | Background initial load must disable unavailable actions and avoid stale DB use; **DERIVED CACHE**, no source/schema change |
| P10 | Settings persistence | `config.py:204` | Whole encrypted settings are rewritten for small usage/cache changes; fixture save can exceed frame budget. Existing temp+replace lacks explicit flush/fsync. | 46 ms median / 57 ms max | Medium | Medium | **STORAGE/ENCRYPTION: approval required.** Do not silently debounce critical saves; any revised path temp+flush+fsync+replace |
| P11 | General search | `drug_db.py:722` | Combined `search_norm` has no index in the cache schema, so its ordered search cannot use the brand/scientific prefix indexes. Recorded EXPLAIN confirms indexed brand prefix versus substring scan; general-search scan/sort is inferred from schema/query inspection. | 10–30 ms at 22k rows | Low | Small–Medium | **SEARCH LOGIC/CACHE:** approval before changes; preserve normalization, prefix-first order and limit. FTS is not a drop-in equivalent |

## Recommended first approval batch

1. P01–P02: reduce redundant/full-size glass work and same-page remapping while keeping the current appearance.
2. P03: batch template insertion; renumber and refresh once.
3. P04 and P08: shorten debounce carefully and use queue-only worker/UI handoff.

P05 offers a major responsiveness win but is deliberately separated because it touches persistence. Start with serialized background writes using captured data and the existing format; do not introduce SQLite patient storage, weaken DPAPI, or silently discard unsaved writes. A decrypted cache would require explicit invalidation/concurrency tests and separate approval.

## Already working well — retain

- Indexed/normalized background brand and scientific search, stale-result protections, mapping-list virtualization and existing favorites/template limits.
- Background snapshot-based cloud/document generation; normal TLS and short-link QR contract.
- Lazy python-docx import, process-level font registration and action-icon caches.
- Explicitly no unfinished-prescription restoration; collapsed template/reference sections and preserved browsing position.

## Verification gates for any approved fix

- One logical change at a time, repeat the matching benchmark sequentially, rerun regressions.
- Search changes: exact dictionaries and ordering must match the baseline for English/Arabic/partial/misspelled queries.
- Persistence: existing encrypted files load unchanged; atomic writes, serialized updates, failure/shutdown recovery, external modification and key privacy verified.
- UI batching: all rows and medication values retained, no Tk calls from worker threads, stale callbacks ignored after close.
- Rendering/export: visual comparison in English/Arabic and A4/A5, same document content, no change to Word layout/calibration or cloud payload.
- Do not claim targets met from constructor/first-idle timings, Python-only memory samples, or window-discovery packaging measurements.

## Limitations and pending checks

The synthetic UI run is exploratory; startup and first-page timings overlapped an initial core benchmark and debug tracing adds overhead. Final core timings were rerun alone. There is no OS-reboot cold-start, native-memory/RSS profile, slow-disk test, real network test or cross-Windows test in this audit. Packaging comparison is a local window-discovery proxy, not a release packaging change. Further measurements are needed before claiming a universal startup budget.

**Approval boundary:** choose the first fix batch and separately approve any storage/encryption/search changes. Version stays 7.2 during audit; the next approved functional update should increment it.
