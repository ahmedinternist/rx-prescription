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
- The application does **not** store patient prescriptions or history. It stores
  clinic/doctor preferences and its signing key using Windows DPAPI, tied to the
  current Windows account.
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
- **Bilingual UI + documents** — switch English / Arabic from the toolbar; the
  app, the PDF, the Word file, and Arabic text shaping all follow.
- **Prescriber + Patient** shown **side by side** in a compact layout (no fixed
  lengths). Prescriber: name, license No., specialty. Patient: name, age, sex.
- **Medications** — repeatable rows. Each row shows the **drug name on top**
  with **Dosage / Frequency / Duration / Notes** boxes directly beneath it.
  Autocomplete from an importable drug database; no fixed drug limit.
- **Paper size**: A5 / A4 / Letter (whole layout adapts).
- **Outputs**:
  - **Preview / Print** — full prescription PDF (prescriber, patient, drug
    table, QR).
  - **Medication Label** — a compact **Word (.docx)** file containing the **full
    medication content** (drug, dosage, frequency, duration, notes) and the QR
    code placed a few lines **above the bottom-right corner** (not flush in the
    corner). The QR still encodes the FULL prescription, so scanning it shows
    prescriber/patient/date/Rx too.
  - **Export Word** — fully editable `.docx` of the full prescription.
- **Arabic** renders correctly (RTL shaping via arabic-reshaper + python-bidi;
  Tahoma/Arial font on Windows).

## Drug database (autocomplete + import)
- Ships with `data/drugs.csv` (27 common drugs); copied to AppData on first run.
- **Import / Replace**: `Import DB…` → any CSV **or Excel (.xls / .xlsx)**. REQUIRED
  column is the drug name (header `generic_name`, `generic`, `scientific_name`, `inn`,
  or `name`). Optional: `brand_name`, `strength`, `form`, `category`, `notes`.
- **Merge**: choose "No" to keep existing drugs and add new ones.
- **Export**: `Export DB…` saves the current DB to CSV.

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
