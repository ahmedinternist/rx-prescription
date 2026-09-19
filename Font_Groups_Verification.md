# Independent display fonts — 4.82.0

## Approved scope

Three independent controls in Settings → General → Display fonts, with every
integer **18–56** available. Existing compact white/gray Settings styling is
retained; the Refero typography craft guidance informed consistent role-based
sizes and neutral system-font usage, not a new palette or font family.

- Medication suggestions: default 20; autocomplete and main-search result lists.
- Dose instructions: default 18; dosage/frequency/duration/notes fields and
  frequency/notes dropdowns in Medication Entry, Favorites and Templates.
- Patient & doctor names: default 18; name fields and patient-name suggestion list.

Settings menus, titles, action buttons, category/sort filters and context menus
retain their original fonts. Only the explicit live sample labels in Settings
reflect the selected display size. Word/PDF typography and cloud payloads are
unchanged. Older preferences migrate independently; sizes below 18 are raised to
18. Existing encrypted records/configuration are not rewritten merely by opening
the app; saving Settings persists the three new preference keys.

## Verification

- Isolated regression suite: **155 passed**; unsafe legacy `test_headless.py`
  remains excluded. No production patient records are used in tests.
- Tests verify range validation, independent migration, encrypted persistence,
  selected instruction text and popup size consistency, updated/new rows,
  preserved Arabic/mixed name input, unchanged neutral menus and fixed Settings
  dropdown fonts at maximum size 56.
- Isolated Settings screenshots visually checked:
  `output/font-groups-en-True.png` (all groups 56),
  `output/font-groups-ar-False.png` (28/18/22).
- Entry height adapts to selected role font; popups use font-derived row metrics.
  Large text in narrow editable fields can be horizontally scrolled rather than
  changing the stored value or reducing the requested font size.
- Separate executable is packaged in `dist/FontGroups/RxPrescription-v4.82.exe`,
  retaining version 4.82.0 and previous build directories. Frozen startup and
  physical printing are not part of this display-font QA.
  Archive inspection verified all three role keys and the tagged linked trade-name
  picker, as well as the preserved cloud export workflow.
  Size: **42,096,785 bytes**.
  SHA-256: `6e97e5d916de58a47cab83eb1ad8b1c4578885ecd25e621a68ddd1eff04172bf`.
