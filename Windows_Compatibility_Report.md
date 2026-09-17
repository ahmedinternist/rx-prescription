# Windows compatibility audit — v4.82

Test date: 2026-09-17. This is a local Windows 11 audit, not certification of
every Windows release. No Windows 10/8.1/Server/ARM virtual machines were tested.
The original compatibility audit below was performed before the glass-dashboard
styling update. Its executable size/hash describe that earlier snapshot, not
the rebuilt executable. After the glass refinement, 75 automated tests and
136 hidden English/Arabic page checks passed; native glass
APIs were not added (the effect uses portable static textures). Cross-version
Windows compatibility still requires actual machines/VMs; it is not certified.

## Verified environment

- Host: Windows 11, version 10.0.26200, x64.
- Interpreter: CPython 3.13.7, AMD64; Tcl/Tk 8.6.15.
- Screen reported by Tk: 1280 × 800.
- Executable: `dist/RxPrescription-v4.82.exe`, 42,053,779 bytes.
- SHA-256: `e4e5d09cb01ba9cb886f1afe2477824bcb1ffa45015b8ddd14ddc34ecb9be21a`.
- Installed versions: CustomTkinter 6.0.0, Pillow 12.3.0,
  cryptography 50.0.1, ReportLab 5.0.1, python-docx 1.2.0,
  google-genai 2.22.0, PyInstaller 6.22.2.

## Results

| Check | Result |
|---|---|
| Standard pytest suite | 73 passed |
| Installed dependency consistency (`pip check`) | No broken requirements |
| Dashboard and Settings UI probe | Passed |
| Treatment-template UI/workflow probe | Passed |
| Online-reference cards/full-label probe | Passed |
| Compact layout, medication tools, class navigation, dropdown sizing | Passed |
| All 8 main pages and 9 Settings sections, English/Arabic, at 100%, 125%, 150%, 200% widget scaling | 136 hidden page/section checks passed in 8 fresh launches; no callback exceptions |
| Windows DPAPI encryption/decryption with Arabic text | Passed |
| Segoe UI, Tahoma, Arial availability on this host | Present |
| Executable PE inspection | AMD64/x64, GUI subsystem |
| Bundled native library PE inspection | All 95 DLL/PYD files parsed successfully; all AMD64/x64 |
| Essential packaged resources | Drug seed CSV, edit icon, viewer HTML, Python 3.13 DLL and both Visual C++ runtime DLLs present |

These UI checks instantiate the source application with hidden windows.
They test callback/workflow behavior, not pixel-perfect visibility at every DPI,
screen size, or Windows theme. The packaged executable was inspected, not
launched as a frozen GUI during this audit. Live Gemini/OpenFDA requests, printer
drivers, Word installation, Defender/SmartScreen and installer behavior were
not certified by these checks.

An older local UI probe initially expected the previous blue navigation
selection rather than the current teal selection. Its test expectation was
corrected and the probe passed. A first multi-launch harness reused Tk roots
inside one process and hit stale interpreter-bound font/image handles. The
harness now tests separate processes, matching normal application launches;
this synthetic harness failure is not evidence of a Windows-version failure.

## Windows-version matrix

| Windows target | Status for this executable |
|---|---|
| Windows 11 x64, build 26200 | Source tests/UI checks passed on this host; frozen startup and full interactive acceptance still unverified |
| Other Windows 11 x64 builds | Not tested |
| Windows 10 x64 | Candidate deployment target; not tested on Windows 10 |
| Windows 8.1 x64 | Meets Python's documented OS baseline only; complete dependency stack and frozen startup not verified |
| Windows 8.0, 7, Vista, XP and older | Below Python 3.13's documented baseline; current build should not be advertised as compatible |
| Any 32-bit Windows | Incompatible with this x64 executable and its x64 native libraries |
| Windows on ARM64 | No native ARM64 build supplied; x64 emulation behavior not tested |
| Windows Server editions | Not tested; evaluate desktop/GUI availability, architecture and dependency requirements separately |

Python 3.13 documentation states Windows 8.1 and newer as its interpreter
baseline. This is **not** a guarantee that the current application plus every
third-party binary runs on Windows 8.1.
[Official Python Windows documentation](https://docs.python.org/3.13/using/windows.html).
Saved evidence: `.firecrawl/python-313-windows.md`.
Pillow 12 supports Python 3.13 according to its
[Python support table](https://pillow.readthedocs.io/en/stable/installation/python-support.html);
that table does not certify every Windows version.
Saved evidence: `.firecrawl/pillow-platform-support.md`.

## Compatibility risks requiring actual target-machine tests

1. The current GitHub workflow uses only `windows-latest` with Python 3.13.
   It does not establish compatibility with older consumer Windows releases.
2. Requirements have broad minimum versions. A future rebuild can select
   different packages and change compatibility even if the source is unchanged.
3. Arabic PDF font discovery uses fixed `C:/Windows/Fonts` paths, with a
   Helvetica fallback. A nonstandard Windows folder or missing fonts can
   produce missing Arabic glyphs. This host has the expected fonts.
4. Settings/data encryption uses the current Windows account's DPAPI keys.
   Copying encrypted settings to another account or machine is not equivalent
   to an OS compatibility test and may prevent decryption.
5. One-file extraction requires a writable temporary directory. Locked-down
   PCs, antivirus rules and standard-user permissions need frozen-build tests.
6. Large fonts and fixed minimum window sizes need visible layout checks on
   smaller screens and high-DPI displays; hidden callback checks cannot prove
   that every control fits within the viewport.

## Remaining real-machine/VM acceptance tests

Use a separate disposable profile on each target; never use real patient data.
At minimum, test Windows 10 x64 and more Windows 11 builds. Test Windows 8.1
only if legacy deployment is genuinely required; keep any obsolete test OS
isolated and do not store patient records there.

For each target, record exact edition, build, architecture, account type,
screen resolution and DPI. Then test:

- Frozen EXE startup with no Python installed, under a standard non-admin user.
- Cold startup, close confirmation, repeated restart and persisted settings.
- Every page/subpage, icon action, medication reorder, dropdown and keyboard selection.
- Drug/template CSV, XLSX and XLS import/export, replacement and validation.
- Patient/history, favorites, mappings, templates and recovery workflows.
- Arabic/English text entry, search, saving and reopening.
- Word exports with/without header and Arabic PDF glyphs; opening exported files.
- 100%, 125%, 150%, 200% DPI, small screens, multiple monitors and monitor changes.
- Offline startup, network failures, invalid API key and successful live lookups.
- Temporary-directory restrictions, antivirus/SmartScreen and non-ASCII user paths.
- Backup/restore according to the intended account/machine migration policy.

To repeat the source checks here, use PowerShell from the project directory:

```powershell
$env:RX_APP_DATA_DIR = 'E:\Prescription App\build\windows-compatibility-check-data'
$env:PYTHONPATH = 'E:\Prescription App'
& 'E:\PDF\python.exe' -m pytest -q
& 'E:\PDF\python.exe' 'build\windows_ui_matrix_probe.py'
& 'E:\PDF\python.exe' 'build\dashboard_icons_probe.py'
& 'E:\PDF\python.exe' 'build\treatment_views_probe.py'
& 'E:\PDF\python.exe' 'build\openfda_cards_probe.py'
& 'E:\PDF\python.exe' 'build\compact_layout_probe.py'
```

Adjust paths on another machine. These commands require the source tree,
the probe scripts under `build`, Python with Tcl/Tk, pytest and the application
dependencies. The environment variable isolates test data from the normal
application profile. Passing them on one OS proves only that tested environment.
