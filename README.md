# Prescription Printer with QR Code

A bilingual (English / Arabic) desktop app for doctors to print prescriptions
(PDF or editable Word) that include a **QR code** carrying the full
prescription. When a pharmacist (or anyone) scans it with a phone camera +
internet, they land on a free static page that shows the prescription — no
server, no backend; the data stays inside the QR code itself.

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
- Patient records and explicitly saved prescription snapshots remain local and
  are protected with Windows DPAPI for the current Windows account. They are
  never uploaded by the app.
- Prescription fields are validated before export. Duplicate medicines and
  overly dense QR codes require an explicit confirmation.
- New QR payloads are signed with an ES256 clinic-local key. To display
  **Verified** in the static viewer, copy the public key from Settings and add
  it to `TRUSTED_SIGNERS` in the deployed `viewer.html`. Until that step, a
  signature is correctly shown as unregistered rather than verified.
- A QR embeds the prescription data. Anyone who obtains the printed QR can
  decode it; it is not suitable for confidential record storage or revocation.
- This project is a document-generation tool, not clinical decision support.
  Drug interactions, contraindications, and local prescribing rules must be
  supplied by an approved clinical data source and governance process.

## Features
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
- **Paper size**: A5 / A4 / Letter (whole layout adapts).
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
- Format: `base64url(zlib(json))` shaped as `<viewer_base_url>#<payload>`.
  High error-correction (H). The compact label's QR is identical to the full
  prescription's QR — it just looks different on paper.
- `viewer.html` is the static page a pharmacist opens. Host it free on GitHub
  Pages / Netlify and set its URL in **Settings**. The page never sends data to
  a server — it decodes the fragment after `#` in the browser.

## Files
- `config.py`         – settings/persistence (paper, language, viewer URL, DB path, profile)
- `i18n.py`           – English / Arabic strings
- `drug_db.py`        – CSV drug database: load / import(replace|merge) / export / search
- `qr_utils.py`       – prescription model + QR encode/decode
- `pdf_generator.py`  – A5/A4/Letter full PDF + compact label PDF + Word export (Arabic-aware)
- `main.py`           – CustomTkinter desktop GUI
- `viewer.html`       – static decoder page for pharmacists
- `data/drugs.csv`    – seed drug database
