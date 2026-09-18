# Cloud prescription QR verification — 4.82.0

## Desktop implementation

- Explicit mobile-viewer payload from `Prescription.to_cloud_payload()`; UTF-8
  JSON POST with a captured encrypted key. Historical `to_qr_payload()` remains
  available for old QR compatibility, not active cloud exports.
- `doctor`, `patient`, `date`, `medications`; entered license, clinic phone and
  patient age are optional. Medication fields are `tradeName`, `genericName`,
  `dosage`, `instructions` (frequency plus notes), `duration` and `quantity`.
  No local patient identifiers, full records, logo paths, keys or coordinates.
- 15-second timeout, normal TLS verification, no automatic retries or redirects.
- Validated HTTPS cloud `/p/` URL passed directly to the QR generator.
- Worker upload, QR and document generation; UI-thread callback queue.
- Captured prescription, paper/language/document settings survive edits and retries.
- Explicit Retry / Export without QR / Cancel; no inline fallback.
- Static decoder excluded from both packaging entry points; legacy printed QR support
  remains dependent on the previously hosted viewer, not this desktop build.
- Existing encrypted records/settings and previous executable builds retained.

## Verification performed

- `E:\PDF\python.exe -m pytest -q`: **119 passed**. Networking mocked, application
  data isolated using `RX_APP_DATA_DIR`. Unsafe legacy `test_headless.py` excluded.
- Tests cover successful UTF-8/Arabic uploads, optional/brand-only payloads, headers,
  missing/rejected keys, connection/timeout/rate-limit/server errors, redirects,
  malformed responses/URLs, encrypted persistence and safe error output.
- Retry/without-QR/cancel handlers, repeated-click guards, snapshot consistency,
  worker callbacks, actual Tk responsiveness and application closure tested.
- A4 and A5 PDF and both Word paths generated with/without QR; document sizes,
  QR image/caption presence and compact Word trailing spacing verified structurally.
- All four PDF pages rasterized and visually inspected. Settings/error modal captured
  with isolated fictitious data and visually inspected.
- Word visual rendering **not verified**: the bundled environment has no LibreOffice
  (`soffice.exe`). OOXML structure checks do not substitute for Word print-layout QA.
- Backend TypeScript normalizer checked using Node 24: minimal v4, Arabic, brand-only,
  scientific-only, legacy text schema and malformed-record handling passed.
- The supplied key was saved only in Windows-encrypted local Settings; no default
  credential was added to the source or executable.
- Cloud test logs contained no test credentials or prescription-body fields.
- Separate executable: `dist/CloudMobileSchema/RxPrescription-v4.82.exe`
  (42,076,729 bytes).
  Archive inspection confirmed v4.82.0, cloud client/async exports, no packaged
  static decoder and no embedded cloud credential. Prior builds remain intact.
  SHA-256: `f70300583212fd7c07f115f1638c587152672d292876700a7efa0a5f36af0a72`.
  Frozen GUI startup and physical printer-driver testing were not performed.

## Live integration verification

One new clearly fictitious mobile-schema prescription was uploaded to the deployed
service, without reading or uploading real patient records. POST and QR creation
succeeded. ZXing decoded the generated PNG and recovered the exact returned short
link: `https://rx-v2.vercel.app/p/6c8cf6b1`. The viewer displayed the entered names,
age, license, phone and medication details, including Arabic instructions,
brand-only and scientific-only items. QR: `output/mobile-cloud-fictitious-qr.png`.

The earlier v4 test initially returned HTTP 500. Its compatibility correction was
subsequently deployed, and the viewer was later redesigned. The historical patch
in `backend-compat/` must not be applied over that new mobile layout.

Companion viewer correction committed to GitHub main as
`6e347f7b92c400fab2a9843c756ba3b5f737a67a`: retain legacy v4 aliases, remove invented
phone/license/location/dosing defaults, and hide call/map actions unless supplied
values are valid. Production deployment verification is pending: Vercel validated
the exact commit on main and accepted the Deploy to Production click, but its
submission remained loading and no new deployment appeared in a separately
loaded deployment list. Production still served `db70bb5` and its invented
defaults at the final check. The pending deployment tab was left open for handoff.
Do not treat the companion viewer correction as live until this commit is Ready
and the fictitious record shows `Not provided` instead of inferred dosing, with
no phone/map actions for its non-diallable test phone and absent coordinates.

Backend seven-day expiry/access policy is unchanged. Link access is not clinician
signature verification. A shared desktop key is protected at rest, not inaccessible
to its Windows user.
