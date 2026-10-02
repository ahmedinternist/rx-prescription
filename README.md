# Prescription Printer with QR Code

A bilingual (English / Arabic) desktop app for doctors to print prescriptions
(PDF or editable Word) with a **QR code containing a short cloud link**.
Minimal prescription JSON is stored by the deployed Next.js/Redis service at
https://rx-v2.vercel.app; scanning opens its `/p/[id]` viewer. Internet is required.

## Run
```
pip install -r requirements.txt
python main.py
```

## Build the Windows executable
```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```
The output name follows `APP_VERSION` (for example,
`dist/RxPrescription-v8.0.exe`). Building requires a full Windows
Python installation with Tcl/Tk; the script stops early if Tk cannot open.

## Safety, privacy, and QR verification
- Saved patient records and history remain local and Windows-DPAPI protected.
  QR exports upload prescriber name/specialty and license, clinic phone, patient
  name and age when entered, date and medication details—not patient sex, IDs,
  allergies, clinic logo paths or signing keys. Anyone holding the link can view
  these fields. Clinic coordinates and website are uploaded only when explicitly
  configured in Clinic Identity.
- Prescription fields are validated before export. Duplicate-medication warnings
  require an explicit confirmation.
- Cloud links are not locally ES256-signed or claimed as signature-verified.
  Anyone holding a printed QR/short link may view the uploaded prescription.
  New links are retained by the deployed backend for 60 days; links created
  before the version 5.7 deployment keep their original expiry.
- This project is a document-generation tool, not clinical decision support.
  Drug interactions, contraindications, and local prescribing rules must be
  supplied by an approved clinical data source and governance process.

## Features
- **Portable glass-inspired dashboard** — a white/light-gray gradient, frosted
  sidebar, search bar, medication/section cards and export footer, with opaque white
  editable fields and charcoal text. Saturated-blue actions and blue navigation selections
  remain consistent; warning/delete colors stay semantic. Glass is simulated
  with static, cached textures rather than Windows-only acrylic or real window
  transparency. Panels sample their position in a shared white/cool-gray backdrop,
  with single hairline borders and restrained white top-edge highlights.
  White-gradient reflections add a top-edge glint and lower diffuse haze instead
  of solid panel fills; the shared backdrop has a matching soft white gradient.
  Favorites, templates, class panels, patient-history cards, reference cards and Settings use the
  same surfaces. Selected-card colors remain visible. A bounded interpreter-local
  weak cache includes position, dimensions and tint; hidden/off-screen panels
  release their images. Scrolling and resizing are debounced, with no live blur
  or animation loop. Search focus changes its outline without shifting geometry.
  This styling does not change prescriptions, exports or saved data (v5).
- **Refined compact controls** — favorite cards retain their natural height;
  brand inputs are bold, scientific inputs regular, and field labels subdued.
  Dashboard and Settings share 24 px rounded line icons, blue when inactive and
  teal when selected, with a pale teal selection and fixed slim indicator.
  Settings keeps its existing ungrouped section order, uses compact navigation
  rows with stationary hover feedback, and neutral borderless content cards.
  Medicine autocomplete popups fit the screen above/below their field, reserve only the needed
  result rows, and scroll for longer lists. Existing keyboard selection remains.
- **Medication tools subpages** — Starred Drugs and Word Preview open in
  separate views with Back navigation, preserving the medication form and its
  values. All four medication toolbar buttons use outlined surface styling.
  The main right-edge scrollbar is hidden on Prescriber, Patient, Treatment
  Template, Interaction Review and Online Drug Reference; wheel scrolling is
  retained. The latter three pages have additional working-card top spacing.
- Starred medicine cards use two columns and a single name/instructions line,
  with a borderless plus action. Long lines use an ellipsis; adding the medicine
  preserves its complete value. Favorite cards use the same plus action.
- Tooltips are removed from every page and subpage. The main search dropdown
  aligns with the full search bar, limits width rather than shifting sideways,
  and uses 21-point result text (5% larger). Settings omits the repeated label
  below its search, retains the native Settings window title, and adds 32 px
  of white bottom padding beneath the footer buttons. Settings retains two distinct reset scopes
  (this page/all settings) and Save Changes; closing with X/Escape uses the same
  unsaved-change confirmation as the removed redundant Cancel button. The
  native Windows minimize/maximize/close title-bar controls remain unchanged.
- **Word with Header** uses a compact letterhead layout based on the supplied
  clinic document: centred clinic/prescriber identity, compact left-aligned
  license/patient/date details, numbered medicine lines, a movable borderless
  logo at the upper-left and a movable borderless QR at the lower-left. It omits
  the prescription title/subtitle, QR caption and closing signature block.
  The separate Export to Word without Header output is unchanged.
- **Compact workspace** — 10 px card padding, tighter card gaps, 36 px input
  fields, and compact action toolbars. The dashboard chevron switches to a
  62 px icon-only rail with tooltips; its state is remembered. Repeated section
  titles and decorative bars are removed. Medication preview shares the top
  toolbar and reserves no space when closed. Clinical warnings remain visible.
- **Bilingual UI + documents** — switch English / Arabic from the toolbar; the
  app, the PDF, the Word file, and Arabic text shaping all follow.
- **Prescriber + Patient** shown **side by side** in a compact layout (no fixed
  lengths). Prescriber: name, license No., specialty. Patient: name, age, sex.
- **Medications** — compact numbered rows: plain number, generic/trade-name
  box, scientific-name box, then movement and delete icons. Dosage, Frequency,
  Duration, Notes, and Quantity remain underneath. Add Drug, Starred Drugs, and
  Save Prescription for Patient share the toolbar below the page title.
- **Shared edit artwork** — `data/edit-icon.png` is the supplied pencil-and-square
  image, processed with the built-in image editor. Prompt: remove only the white
  background (including gaps) to transparency; preserve the black silhouette,
  proportions, and crisp edges. It is bundled with v5 and reused at 18–20 px.
  Autocomplete from an importable drug database; no fixed drug limit. Quantity
  is calculated from dose count, frequency, and duration (including compact
  numeric dose/day inputs), and remains editable when clinical judgment is needed.
- **Continuous prescription workflow** — Patient → Medicines → Review → Export
  uses the same compact progress control on the patient and medication pages.
  Patient context remains visible during entry, each stage exposes one primary
  next action, and Review separates structural errors from clinician-review
  warnings. Unfinished prescriptions are not saved or restored on the next launch.
  Older resume snapshots are discarded on startup; saved patient history is unaffected.
- **Compact references and stable browsing** — OpenFDA sections start collapsed;
  opening a section closes the previous section. Adding favorites or treatment
  templates keeps the current browser page and scroll position. Favorite selections
  and the expanded saved-template card remain selected.
- **Progressive medicine rows** — when another medicine is added, completed rows
  collapse into a numbered one-line summary and reopen from the shared edit icon.
  Medication frequency offers editable Arabic presets. Notes occupy the former
  Quantity field space without changing the other field columns. Existing saved
  quantity data remains compatible but no Quantity control is displayed.
- **Drug-class browser** — detailed-class medicine rows support Use in Rx,
  mapping, starring, and guarded deletion from the local drug database.
- **Treatment templates** — compact icon actions create, save,
  and delete reusable treatment plans.
- **Patient details and history** — compact demographics, Arabic-aware entry,
  duplicate detection, saved/modified status, and a collapsible prescription
  timeline. The current medicines are compared with the latest saved prescription,
  highlighting additions, removals, and changes.
- **Recovery centre** — deleted favorites, templates, patients, prescriptions,
  and class mappings can be restored for 30 days from Settings.
- **Paper size**: A5 / A4. Both Word export variants explicitly embed the selected
  page dimensions. Re-export older Word files to update their paper size; printer
  driver settings can still override the document's paper choice.
- **Display fonts**: Settings → General offers dropdown/autocomplete and patient-name
  field font sizes from 10–56 px. Dropdown **Default** preserves each control's original
  size. Settings menus always keep their normal size, independent of this preference.
  These controls change the interface, not the prescription's export typography.
- **Outputs**:
  - **Preview / Print** — full prescription PDF (prescriber, patient, drug
    table, QR).
  - **Export to Word without Header** — a compact editable `.docx`
    containing the numbered medication lines and QR code without the clinic
    header.
  - **Export Word with Header** — a fully editable `.docx` containing
    the clinic header, prescriber and patient details, numbered medication lines,
    and QR code.
- **Arabic** renders correctly (RTL shaping via arabic-reshaper + python-bidi;
  Tahoma/Arial font on Windows).

## Drug database (autocomplete + import)

Treatment Template has two views: **New Template** opens the existing disease
and medicine editor; **Saved Templates** shows searchable A–Z cards, expandable
treatment details, and Add to Rx, Edit, and confirmed Delete actions. Saving
returns to the browser and highlights the saved template. Leaving a changed
editor prompts before discarding changes. Existing template storage, medicine
fields, OR alternatives, recovery, and Settings database import/export remain.
Saved Templates now uses a responsive two-column card library with bold disease
titles, medicine-count badges, visible expand controls, distinct empty/search
states, and sorting by usage, recent activity, modification date, or name. Cards
start collapsed and reveal medicines only when their disease title or chevron is
selected; opening one card closes the previous card. Search is debounced, results
remain batched, and editing or applying a template preserves its query, expanded
item, and scroll position. Existing templates gain usage metadata automatically
without changing their medicines.

Online Drug Reference presents five expandable, color-coded label sections:
Indication (teal), Dose (blue), Contraindications (red), Pregnancy (amber), and
Renal Adjustment (purple), replacing the previous Interaction section. Cards
use independent vertical columns on wide pages to avoid uneven row-height gaps,
and stack in section order on narrow pages. Medicine identity, FDA label,
label/check dates and View Full Label share one compact header row; the excerpt
disclaimer footer is omitted. Missing information
is explicitly marked as not stated. Indications and dose come from the label's
corresponding fields; renal excerpts come from topic-containing dosage,
population, precaution, or contraindication sections, not inferred dose rules.
Expand a card for full extracted context or use View Full Label; cards do not
include a bottom Open label source link. The page has no introductory subtitle.
See the [official openFDA label fields](https://open.fda.gov/apis/drug/label/searchable-fields/).
The page also supports a medicine search independent of Medication Entry,
searchable full-label viewing with highlighted matches and larger bold section
24-point headings, excluding Pediatric Use, How Supplied, Warnings and Cautions,
Use in Specific Populations, Clinical Studies and Clinical Pharmacology,
section-source labels,
and effective-date warnings. Dose displays explicit label dosage plus separately
identified adult, maximum, route, and hepatic statements when those are actually
present; it does not add a pediatric-dose section. Renal information is marked
as a stated adjustment, precaution-only text, or not found. Successful results
are cached for seven days inside the app's Windows-encrypted settings, visibly
marked when reused, and can be refreshed or cleared from the page.

The Drug Classes browser places its two columns directly below the search bar,
without group/class dropdown filters, an unclassified count line, or group stars.
Major groups and detailed classes use visible open arrows rather than hidden
double-click navigation. A major-group page lays its detailed classes out in two
columns, while a detailed-class page keeps the breadcrumb, medicine count, and
borderless **+** action together in one compact header (with no second search box).
Back navigation restores the previous group, detailed class, query, and scroll
position. Class Mapping Editor uses compact All, Unclassified, Suggested, and
Conflicting filter chips with an inline result count. Its queue font is independent
of global autocomplete sizing, and Save & next retains the review position.

Treatment Template appears in Dashboard → Reusable Content directly below
Favorite Drugs; Drug Classes remains under Clinical Reference.

Dashboard navigation uses matching 24 px blue line icons, a fixed icon/label gap,
and a pale-blue active-page highlight with a non-shifting indicator. Dashboard
icons and the collapse control have no tooltips, including in collapsed mode.

The approved light Clinical Glass layout uses white outlined Medication Entry
actions and direct up/down arrows beside each medicine. Drag sorting remains
available on its plain row number. Frequency and notes stay editable and use
white fields. Settings groups gray cards on a softly graduated white canvas;
Clinic Identity has three side-by-side contact fields, an inline logo thumbnail
and a live header preview. Glass is simulated with cached static rendering,
not desktop transparency. The current version is 8.0.

Version 8.0 preserves the current patient/medicine form in memory across language
changes, repairs malformed settings sections, waits for exports before closing,
ignores stale OpenFDA lookup results, and bounds heavy browsing card contents.
Off-screen geometry spacers retain scrolling; unfinished prescriptions still
are not saved or reopened after application restart.

Version 7.9 blocks writes to unreadable patient history, keeps same-name patients
separate unless an existing patient ID is explicitly selected, includes encrypted
patient history in manual/automatic backups, recovers invalid derived medicine
caches on startup, and serializes medicine database mutations. Older backups
without patient history leave existing history untouched. DPAPI backups remain
bound to the original Windows account; this is not cross-PC key migration.

Version 7.8 keeps Favorite Drugs cards compact: only medicine names and the
category are displayed, with the category at the end of the action row.
Dosage, frequency, duration and notes remain available in the editor and are
preserved when adding a favorite to the prescription.

Version 7.7 wraps long medicine names within their card columns and reserves
space for action icons. Compact bilingual save feedback distinguishes saving,
successful persistence and failures without changing saved medication values.

Version 7.6 shows values-only regimen summaries on treatment cards while retaining
editor field labels. Medication Entry permits only one expanded editing row at
a time, including incomplete rows. Mapping shows trade names alongside scientific
names with larger text; the selected-medicine heading follows the same order.

Version 7.5 unifies card corners and spacing, uses bold primary names with
regular scientific names and subdued regimen details, and standardizes Add,
Edit, Delete and Star icon hit areas without visible boxes or tooltips. Focus
feedback changes only the field border color and preserves warning/error colors.

Version 7.4 pauses favorite/template card construction beyond the viewport plus
a small buffer, resuming automatically as the user scrolls. Glass redraws wait
200 ms after window resizing settles. Replacement templates reuse existing
medication rows, invalidate old autocomplete callbacks and clear unused rows.

Version 7.3 keeps glass image processing off the Tk thread, clips tall panels to
their visible slice and bounds the image cache. Same-page navigation avoids
remapping, and favorites/templates reuse unchanged cards with incremental initial
rendering. Hidden card construction pauses until its page is shown. Autocomplete
uses a 35 ms debounce without changing matching or result order. Template
application yields between rows and retains its modal selection dialog until done.

Prescription saves use a captured snapshot and a dedicated serialized worker;
repeat clicks are ignored while saving, and newer edits are never marked saved.
Closing through the UI waits for the pending save. Patient history caches decrypted
records with file-change invalidation; returned records cannot mutate that cache.
Settings and history retain the existing DPAPI JSON formats and use flushed atomic
replacement. Unchanged settings writes are skipped only while the backing file
also remains unchanged. No unfinished prescription is restored.

Drug-cache loading is deferred until after window construction, and ReportLab/font
initialization is lazy. The disposable SQLite cache adds a combined-name index,
without changing normalization, prefix-first ordering or substring behavior.
openFDA worker callbacks use the UI queue. See `PERF_REPORT.md` for measured
improvements, remaining budget misses and verification limits.

Version 7.2 builds reference, interaction, favorite, classification and template
pages on first use and reuses their controls. Startup backup compression and
classification indexing run in background workers. Superseded medication
autocomplete jobs are cancelled when still queued. Bounded local timing samples
contain only fixed operation names and durations, never prescription content.

Version 7.1 renders only the mapping-list viewport, saves captured mapping/name
changes in one background CSV write/cache rebuild, and defers patient comparisons
while their page is hidden. Unchanged comparison results reuse their widgets.
While a mapping save is pending, repeat saves and editor-context changes are
blocked; failed saves retain the pending edits for retry.

Visual polish keeps input heights compact while enlarging medication labels;
cached, consistently sized action icons retain their existing commands and the
supplied edit artwork. Neutral entry/dropdown borders turn blue while focused
without resizing or hover motion. Favorites and Template search toolbars share
one height/corner rhythm, and the export footer uses a lighter white sheet.
No tooltips, clinical record changes or document-format changes are introduced.

- Ships with `data/drugs.csv` (27 common drugs); copied to AppData on first run.
- **Import / Replace**: **Settings → Database → Import Database** → any CSV or
  Excel (`.xls` / `.xlsx`). Each medicine row needs either a scientific/generic
  name or a trade/brand name; brand-only products are retained. Supported name
  headers include `generic_name`, `generic`, `scientific_name`, `inn`, `name`,
  and `brand_name`. Optional: `strength`, `form`, `category`,
  `therapeutic_group`, `detailed_class`, and `notes`. Imported group labels and
  detailed classes are normalized to the app taxonomy. Missing detailed classes
  can receive suggested mappings; explicitly unrecognized class values remain
  visible as needing review instead of being converted silently to `Other`.
- The Medication Entry **Generic / trade name** box searches the imported trade
  field (or a legacy row's only name). The **Scientific name** box searches only
  the imported scientific / INN field and can offer linked trade names afterward.
- **Merge**: choose "No" to keep existing drugs and add new ones.
- **Export**: **Settings → Database → Export Database** saves the current database
  as CSV, XLSX, or legacy XLS.

## QR payload & viewer
- **Settings → Clinic Identity → Clinic website URL** accepts a public HTTPS
  clinic website (a bare domain is normalized to HTTPS). When present, it is
  added to new cloud export snapshots and appears as a globe action in the
  scanned viewer. It is not embedded in the QR image or saved prescription.
- **Preview cloud viewer** opens a local mobile-sized preview with fictitious
  prescription data and the current unsaved clinic identity. It creates no
  network request, file or Redis record.
- **Settings → Clinic Identity → Clinic Location** accepts a Google Maps dropped-pin
  link, a supported Google short link, or explicit latitude/longitude. Check and
  confirm the pin before saving. Map-centre (`@…`) coordinates are deliberately
  not treated as the clinic pin. Coordinates are stored in encrypted configuration.
- **Use Current Location** opens a temporary loopback-only page in your default
  browser. Click its location button, grant browser permission, then return to
  confirm the clinic pin. Use it only while physically at the clinic. It does not
  enable Windows location services, bypass permission or track continuously. The
  page expires after two minutes; denied/unavailable location has a manual fallback.
- **Include clinic location in cloud prescriptions** is off by default. Enabling
  it adds validated numeric `latitude` and `longitude` to new export snapshots;
  the cloud viewer's location icon opens that pin in Maps. Anyone holding the link
  can see it. Removing the location also disables inclusion. Changing Settings
  does not modify previously uploaded prescriptions. Printed Word/PDF layout and
  the QR short-link format remain unchanged. Small Clinic Identity views can scroll.
- Open **Settings → QR verification**, enter the cloud API key and save Settings.
  The key is masked and stored in Windows-encrypted configuration, not in the executable.
- QR verification also shows whether the key is configured, whether the cloud
  API/Redis and public viewer are reachable, and the encrypted timestamp of the
  last successful prescription upload. Its connection test performs an authenticated
  read-only health check and does not create a prescription.
- Export snapshots the explicit mobile-viewer JSON contract and POSTs it to
  `/api/rx` on a worker: optional `clinicName`, `doctor`, `registrationId`, `phone`,
  optional `website`, `patient`, `age`,
  `date`, and `medications`. Each medicine uses `tradeName`, `genericName`
  (scientific name), `dosage`, `instructions` (frequency and notes), `duration`,
  and `quantity`. Unentered optional fields are omitted; no dosing is inferred.
  `to_qr_payload()` retains the historical version-4 shape but is not uploaded
  by active exports; `to_cloud_payload()` is the active contract.
  Only the returned HTTPS `/p/[id]` URL enters the QR generator (error correction H).
  Export buttons stay disabled while upload/document creation runs; form edits do
  not affect the captured prescription. POSTs time out after 15 seconds and are
  never automatically retried or redirected.
- If upload fails, choose **Retry**, **Export without QR**, or **Cancel**.
  Export without QR requires an explicit choice and omits QR/caption/QR-only spacing.
  There is no inline-data fallback. A timeout may occur after storage succeeded;
  an explicit retry can create another cloud record.
- Legacy encoding/signing helpers remain available as compatibility APIs but are
  unused by active exports. Retired local sample-viewer files are no longer shipped.
  Old printed QR codes still depend on their original hosted viewer and are not migrated.
- The new mobile-schema fictitious upload, actual QR-image decoding and rendered
  prescription were checked successfully. The historical v4 compatibility patch
  remains recoverable from Git history and must not overwrite the current mobile layout. See
  `Cloud_QR_Verification.md` for the latest viewer deployment status and QA limits.

## Files
- `config.py`         – encrypted settings/persistence, including cloud API key
- `i18n.py`           – English / Arabic strings
- `drug_db.py`        – CSV drug database: load / import(replace|merge) / export / search
- `qr_utils.py`       – prescription models + QR image generator; deprecated legacy helpers
- `cloud_rx.py`       – authenticated cloud link client with safe errors and URL validation
- `clinic_location.py` – clinic pin validation and permission-based loopback browser helper
- `pdf_generator.py`  – A5/A4 full PDF + compact label PDF + Word export (Arabic-aware)
- `main.py`           – CustomTkinter desktop GUI
- `data/drugs.csv`    – seed drug database

## Independent display fonts

Settings → General → Display fonts contains three independent controls, each
with every integer size **18–56** and a live preview:

- Medication suggestions (default 20): medicine autocomplete in Medication Entry,
  Favorites and Templates, plus main-search suggestions and linked trade-name lists.
- Dose instructions (default 18): dosage, frequency, duration and notes fields,
  including frequency/notes dropdowns in Medication Entry, Favorites and Templates.
- Patient & doctor names (default 18): patient and prescriber name fields and
  patient-name suggestions.

Settings menus retain their fixed readable fonts; page titles, buttons, category
filters, sorting menus and context menus do not inherit these overrides. Large
entry fonts get sufficient field height; existing edits are preserved when applying
preferences. These are desktop display sizes and do not change Word/PDF typography.
Older medicine-dropdown and patient-name preferences migrate to their corresponding
groups; values below 18 are raised to 18, and dose instructions start independently
at 18. Defaults or invalid old values use 20/18/18. Encrypted records are preserved.
