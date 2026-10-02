# Performance baseline — Rx Prescription Printer 7.2

Audit date: 2026-10-01. Phases 1–2 only; no performance fixes or production packaging changes applied.

## Detected stack and architecture

- Windows desktop: Python 3.13.7, Tkinter/Tcl/Tk and CustomTkinter. `main.py` owns the UI, medication rows, navigation and glass rendering. There is no desktop Electron, .NET or browser UI.
- Drug source: CSV; XLS/XLSX imports are converted into this source. `drug_db.py` maintains a derived SQLite cache with normalized columns and prefix indexes. Brand/scientific autocomplete runs in background workers.
- Settings/favorites/templates: in-memory configuration persisted as Windows DPAPI-encrypted JSON. Patient history: a separate whole-file DPAPI-encrypted JSON document, decrypted and parsed afresh by history operations. No database-format changes were made.
- PDF: ReportLab, Pillow, Arabic reshaping/bidi. Word: python-docx, already imported lazily in export functions. Font registration occurs once at PDF module import.
- Cloud QR: HTTPS client/short link; document exports use captured snapshots and background workers. No real cloud upload was performed during this audit.
- Concurrency: three-worker executor plus UI-thread queue polling; some openFDA daemon-worker callbacks still invoke Tk `after` directly.
- Packaging: PyInstaller onefile/windowed; an isolated onedir diagnostic prototype was built solely for comparison. Earlier executables are retained.

## Method and safety

`perf_baseline.py` creates only fictitious data under `output/perf-audit-v72/isolated-data`, selected through `RX_APP_DATA_DIR`. It exercises real application methods without network requests. Real AppData, API keys, prescriptions and existing executables were not modified. Generated English/Arabic A4/A5 PDF and both Word outputs contain only synthetic information and no QR.

Fixture: 22,000 drugs, 550 templates containing ten medicines each, 600 favorites, 2,200 patients and initially 6,297 prescriptions. One patient has 300 historical prescriptions. Encrypted history starts at 10,909,222 bytes; configuration at 2,021,158 bytes. Save benchmarks append a few additional fictitious prescriptions.

Optional `RX_PERF_DEBUG=1` instrumentation in `perf_probe.py` records fixed operation names and durations only, never arguments or clinical content. Default application runs do not install wrappers. Samples are bounded. Removing the module and the two guarded hooks in `main.py` removes instrumentation.

Core measurements below were repeated sequentially after the UI run. The first core/UI runs overlapped: first-page/startup UI numbers are exploratory and may include contention. UI tracing also adds overhead. These are local synthetic timings, not controlled reboot-cold or cross-Windows benchmarks. First `app.update()` includes painting and pending callbacks; it is not equivalent to constructor time. No visual output equivalence or real printer validation is claimed.

## Measurements

| Operation | Observed time | Target / interpretation |
|---|---:|---|
| Import `main` | 529 ms cumulative | ReportLab/PDF module 162 ms; CustomTkinter 113 ms |
| Warm drug-cache load | 0.84 ms | Already efficient |
| Cold derived drug-cache rebuild | 1,649 ms | Synchronous startup path; separate from import time |
| Indexed brand autocomplete query | 0.23–3.87 ms medians | Do not replace search unnecessarily |
| Indexed scientific autocomplete query | 0.56–5.63 ms medians | Already efficient |
| General combined-name search | 10.2–30.0 ms medians | SQLite scan/temp sort; still below 100 ms |
| Last key to visible autocomplete | 203–238 ms | Includes 180 ms debounce, worker and UI polling; target <100 ms |
| First constructor-to-update | 3,893 ms | Exploratory, excludes imports; not a rigorous cold-launch result |
| First favorites / templates / classes page | 11,020 / 4,388 / 3,243 ms | Exploratory; includes layout and glass paints |
| Favorites refresh | 838 ms median | Reuses cards, but updates many widgets synchronously |
| Saved-template refresh | 710 ms median | Rebuilds visible collapsed cards |
| Populate 2,200-patient list | 115 ms median | UI batch exceeds 50 ms |
| Display 300 prescription-history cards | 1,395 ms | Includes encrypted load and widget construction |
| Add populated medication row | 147 ms median | Target <100 ms |
| Apply ten-drug template | 1,976 ms | Target <200 ms |
| Medication change / six-row preview | 0.52 / 37.8 ms medians | Do not prioritize change-handler micro-optimizations |
| History decrypt + parse | 898 ms median | Whole-file DPAPI + JSON read |
| History get / search-all | 726 / 714 ms medians | Repeated whole-file loads |
| History save, service only | 1,003 ms median | Synthetic repeated run; variation 497–1,015 ms |
| Save prescription, whole UI handler | 1,751 ms | Noticeable main-thread blocking |
| Configuration encrypt/write | 46 ms median, 57 ms maximum | Approaches/exceeds UI budget at fixture size |
| Full classification index | 2,430 ms median | Already backgrounded; consumes a shared worker |
| Ten scripted resize/update steps | 16,065 ms median | Not a physical drag; repeated paints dominate |
| Individual glass paint | Up to 2,156 ms | Large full-size raster operation on UI thread |
| Local PDF export, four EN/AR A4/A5 cases | Approximately 17–19 ms | No QR/network; not end-to-end cloud timing |
| Word header / no-header generation | Up to 140 / 40 ms | Local generation only; exports already backgrounded |

Forty page switches retained approximately 154 KB according to `tracemalloc` (peak approximately 467 KB during that interval). This does **not** measure native Tk/Pillow allocations, process RSS, or establish absence of leaks. No UI callback exceptions occurred in the scripted run.

## Search and regression protection

Thirty-nine cases across brand, scientific and combined search were compared with an independent enumeration/sort of the full fixture: English and Arabic one/two/three-character prefixes, partial values, case/whitespace normalization, scientific names and a misspelling. Returned dictionaries and ordering matched exactly. This protects existing behavior, not clinically correct identification. SQLite EXPLAIN output is recorded in `core.json`.

Complete maintained regression suite: **192 passed in 67.33 seconds**, using a separate synthetic AppData directory and test temporary directory.

## Reproduction and evidence

Use `E:\PDF\python.exe perf_baseline.py seed --output output/perf-audit-v72`, then `core`, `imports` and `ui` sequentially. Set `RX_PERF_DEBUG=1` only for the UI timing run. Do not run core and UI against the same fixture concurrently. `perf_package_probe.ps1` compares local packaged window-discovery latency using isolated AppData; it stops only the processes it launched. Window discovery is a coarse proxy, not proof of usable-page readiness.

### Local packaging comparison

Existing v7.2 onefile: window detected in 4,199 and 5,174 ms; first attempt exceeded the 30-second discovery timeout. Diagnostic onedir: 6,282, 5,767 and 3,695 ms. These sequential, warm-machine samples used hidden process launch and process/window polling, whose latency and window-visibility behavior affect detection. The two binaries also differ by the disabled diagnostic hooks. They do not establish a reliable onedir speedup or distinguish extraction, antivirus and painting costs. Do not switch packaging on this evidence. A controlled visible-window bootloader/extraction trace remains necessary.

Raw evidence: `output/perf-audit-v72/{seed,core,imports,ui,packaging}.json` and `importtime.txt`. No claim of controlled cold-start packaging improvement is made.
