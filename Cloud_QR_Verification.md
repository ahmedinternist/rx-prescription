# Cloud prescription QR verification — 4.82.0

## Desktop implementation

- Explicit mobile-viewer payload from `Prescription.to_cloud_payload()`; UTF-8
  JSON POST with a captured encrypted key. Historical `to_qr_payload()` remains
  available for old QR compatibility, not active cloud exports.
- `doctor`, `patient`, `date`, `medications`; entered license, clinic phone and
  patient age are optional. Medication fields are `tradeName`, `genericName`,
  `dosage`, `instructions` (frequency plus notes), `duration` and `quantity`.
  No local patient identifiers, full records, logo paths or keys. Clinic coordinates
  are optional, validated and uploaded only when inclusion is explicitly enabled.
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
values are valid. A subsequent read-only check confirmed this commit Ready at
Vercel deployment `9FG4pXxghthuw3tnQgQWwi1harAu`. Both fictitious records opened
successfully; missing dose values showed `Not provided`, and non-diallable/absent
phone and absent coordinates produced no call/map actions. This records the
previous deployment check, not a new deployment during the clinic-location change.

## Clinic-location addition

- Compact Clinic Identity card with explicit pin links/coordinates, confirmation,
  Maps access, permission-based current-location detection and removal.
- Browser helper binds only to `127.0.0.1`, uses a random one-shot token,
  same-origin/Host checks, bounded bodies, CSP, no-store/no-referrer and no request
  logging. Location permission is requested only after a button click; Settings
  closure cancels the session. No external scripts, geocoder or API key are used.
- Coordinates are Windows-encrypted locally. Inclusion is off by default and
  adds numeric top-level `latitude`/`longitude` only after explicit enablement.
- Automated tests use fictitious pins, mocked browser permission responses and
  isolated Settings. No real geolocation permission prompt or clinic location
  was activated; actual browser/device accuracy still requires manual checking.
- English and Arabic Settings screenshots captured from app-owned test windows:
  `output/clinic-location-settings-en.png`, `output/clinic-location-settings-ar.png`.
- Final isolated regression run: **154 passed**. Added tests cover bounds,
  zero coordinates, explicit versus viewport pins, allowed short-link redirects,
  offline errors, loopback Host/origin/token checks, malformed bodies, one-shot
  responses, simulated permission errors/success, timeout/cancellation, encrypted
  persistence, opt-in payloads, stale callbacks and export snapshot consistency.
- Separate updated executable: `dist/ClinicLocation/RxPrescription-v4.82.exe`,
  **42,093,840 bytes**, retaining version **4.82.0** and all previous builds.
  Archive checks confirmed the location helper, validated-coordinate payload,
  async cloud exports and no bundled static viewer or embedded cloud secret.
  SHA-256: `04b1d0cb5b41868aa7366d6c9733b699e974dd97f72dc0028c88fbbe2ec3ee8d`.
- The loopback helper's secure-context basis was checked against official
  [MDN documentation](https://developer.mozilla.org/en-US/docs/Web/Security/Defenses/Secure_Contexts).
  No new cloud record, backend deployment, frozen GUI launch or physical printer
  test was performed for this addition. The earlier live QR checks remain above.

Backend seven-day expiry/access policy is unchanged. Link access is not clinician
signature verification. A shared desktop key is protected at rest, not inaccessible
to its Windows user.
