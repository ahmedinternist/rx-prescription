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

## Optional Gemini drug reference

The amber **!** beside a populated scientific-name field opens a concise
Gemini Flash reference card for that medicine. The app uses Google's
rolling `gemini-flash-latest` model alias with stable-version fallbacks. It returns
nine validated sections: indications, minimum and usual starting dose, minimum
and usual frequency, maximum dose/frequency, adverse effects, contraindications, pregnancy,
and renal adjustment. Grounded results include returned source links.

1. Create a Gemini API key in Google AI Studio.
2. Open **Settings → Gemini drug reference…**.
3. Paste the key, enable the feature, and use **Test connection**.

The API key and the seven-day lookup cache are encrypted for the current
Windows user. Only the scientific medicine name is included in lookup requests;
patient and prescription data are not sent. Results are reference material for
clinician verification and never fill prescription fields automatically.
The app first requests live Google Search grounding. If Google returns a quota
error, it automatically retries without web search and marks the result clearly
as **Free mode — not web-grounded**; no source links are claimed in that mode.

## Build the Windows executable
```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```
The output is `dist/RxPrescription.exe`. Building requires a full Windows
Python installation with Tcl/Tk; the script stops early if Tk cannot open.

## Safety, privacy, and QR verification
- Saved patient records and history remain local and Windows-DPAPI protected.
  QR exports upload prescriber name/specialty and license, clinic phone, patient
  name and age when entered, date and medication details—not patient sex, IDs,
  allergies, clinic logo paths or signing keys. Anyone holding the link can view
  these fields. No clinic coordinates are currently collected or uploaded.
- Prescription fields are validated before export. Duplicate-medication warnings
  require an explicit confirmation.
- Cloud links are not locally ES256-signed or claimed as signature-verified.
  Anyone holding a printed QR/short link may view the uploaded prescription.
  Retention and link expiry are controlled by the deployed backend.
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
  This styling does not change prescriptions, exports or saved data (v4.82).
- **Refined compact controls** — favorite cards retain their natural height;
  brand inputs are bold, scientific inputs regular, and field labels subdued.
  Dashboard and Settings share 24 px rounded line icons, blue when inactive and
  teal when selected, with a pale teal selection and fixed slim indicator.
  Settings keeps its existing ungrouped section order, uses compact navigation
  rows with stationary hover feedback, and neutral borderless content cards.
  Medicine and quick
  search popups fit the screen above/below their field, reserve only the needed
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
- **Word with Header** retains the clinic logo/name/contact details, followed
  by one prescriber/specialty/license paragraph and one date/prescription-number
  paragraph. It omits the prescription title/subtitle and closing signature
  block. Patient details, numbered medicine lines and QR content remain.
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
  proportions, and crisp edges. It is bundled with v4.82 and reused at 18–20 px.
  Autocomplete from an importable drug database; no fixed drug limit. Quantity
  is calculated from dose count, frequency, and duration (including compact
  numeric dose/day inputs), and remains editable when clinical judgment is needed.
- **Quick prescribing** — press `Ctrl+K` to search patients, treatment templates,
  starred medicines, and drug classes from one keyboard-friendly command box.
  Its large result menu closes after six seconds or when the user clicks outside it.
- **Drug-class browser** — detailed-class medicine rows support Use in Rx,
  mapping, starring, and guarded deletion from the local drug database.
- **Treatment templates** — compact, tooltip-labelled icon actions create, save,
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
  Top search results request twice the previous width and expand for their content,
  while staying aligned with the search bar. Top search width is not capped at the
  screen edge; other autocomplete menus retain their screen-aware width limits.
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
Detailed-class medicine pages use a borderless **+** to add a medicine to the Rx.
Unclassified, suggested, and conflicting mapping review remains in Class Mapping Editor.

Dashboard navigation uses matching 24 px blue line icons, a fixed icon/label gap,
and a pale-blue active-page highlight with a non-shifting indicator. Dashboard
icons and the collapse control have no tooltips, including in collapsed mode.

The approved light Clinical Glass layout uses white outlined Medication Entry
actions and direct up/down arrows beside each medicine. Drag sorting remains
available on its plain row number. Frequency and notes stay editable and use
white fields. Settings groups gray cards on a softly graduated white canvas;
Clinic Identity has three side-by-side contact fields, an inline logo thumbnail
and a live header preview. Glass is simulated with cached static rendering,
not desktop transparency. The current version remains 4.82.0.

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
- Open **Settings → QR verification**, enter the cloud API key and save Settings.
  The key is masked and stored in Windows-encrypted configuration, not in the executable.
- Export snapshots the explicit mobile-viewer JSON contract and POSTs it to
  `/api/rx` on a worker: `doctor`, `registrationId`, `phone`, `patient`, `age`,
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
- `viewer.html` and old encoding/signing helpers are deprecated historical assets,
  excluded from new executable resources and unused by active exports. Old printed
  QR codes still depend on their original hosted viewer and are not migrated.
- The new mobile-schema fictitious upload, actual QR-image decoding and rendered
  prescription were checked successfully. The historical v4 compatibility patch
  in `backend-compat/` must not overwrite the current mobile layout. See
  `Cloud_QR_Verification.md` for the latest viewer deployment status and QA limits.

## Files
- `config.py`         – encrypted settings/persistence, including cloud API key
- `i18n.py`           – English / Arabic strings
- `drug_db.py`        – CSV drug database: load / import(replace|merge) / export / search
- `qr_utils.py`       – prescription models + QR image generator; deprecated legacy helpers
- `cloud_rx.py`       – authenticated cloud link client with safe errors and URL validation
- `pdf_generator.py`  – A5/A4 full PDF + compact label PDF + Word export (Arabic-aware)
- `main.py`           – CustomTkinter desktop GUI
- `viewer.html`       – deprecated historical decoder, not bundled
- `data/drugs.csv`    – seed drug database
