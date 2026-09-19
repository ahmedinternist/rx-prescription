# Backend viewer compatibility fix

## Current mobile viewer (September 2026)

The desktop now uploads the new viewer contract through
`Prescription.to_cloud_payload()`: text `doctor`/`patient`, `date`, optional entered
`registrationId`, `phone`, `age`, and `medications` with `tradeName`, `genericName`,
`dosage`, `instructions`, `duration`, `quantity`. Frequency and notes are combined
as instructions. Missing values are omitted, never inferred. Optional numeric
`latitude`/`longitude` are uploaded only after enabling clinic-location inclusion
in desktop Settings. The viewer already validates their bounds and uses them for
its conditional Maps icon; no backend schema change is required for this feature.

The new mobile layout implements legacy aliases directly. GitHub main commit
`6e347f7b92c400fab2a9843c756ba3b5f737a67a` preserves those aliases and removes invented
clinical/contact defaults. The subsequent read-only Vercel check confirmed this
commit Ready in production; the fictitious viewer also showed the corrected
conditional contact/location actions. See the verification report for evidence.

The new fictitious upload and actual QR-image scan succeeded; its viewer displayed
the entered mobile-schema values at `https://rx-v2.vercel.app/p/6c8cf6b1`.
See `Cloud_QR_Verification.md` for desktop test and packaging evidence.

## Historical patch — do not apply over the new mobile layout

The original viewer accepted text `doctor`/`patient` and `medications`. The earlier
desktop uploaded the minimal v4 shape: object `doctor`/`patient` and `drugs`.
Rendering either object directly as a React child causes the observed HTTP 500.
Missing `medications` would additionally hide the medicines.

1. Copy `rx-normalizer.ts` into the backend as `app/lib/rx-normalizer.ts`.
2. Apply `viewer.patch` from the backend root (or make its exact changes manually).
   The relative import from `app/p/[id]/page.tsx` is `../../lib/rx-normalizer`.
3. Remove the now-unused `Medication` and `PrescriptionData` interfaces from that page.
4. Redeploy the backend and open the already-uploaded fictitious QR. No desktop
   schema change or migration of stored records is needed. Both formats are supported.

Run `node backend-compat/test-normalizer.mjs` in this desktop workspace with Node
24 to repeat the normalization checks. In a backend repository, adjust the test
import to the copied `app/lib/rx-normalizer.ts` location if copying the test too.

This historical patch was subsequently deployed before the mobile redesign. It does not alter
Redis storage, API authentication, access policies or the seven-day expiry.
The wording change avoids claiming digital-signature verification: authentication
of an upload key is not verification of a prescribing clinician's signature.

## Historical verification status

The live fictitious POST returned HTTP 201 and a valid short link. Before this
patch, its viewer returned HTTP 500. The patch addressed that schema mismatch.
The earlier QR image is `output/cloud-fictitious-qr.png` locally. See the current
verification report above for the new mobile schema.

Validate the API's Upstash SET JSON `result`, not only its HTTP status, separately
if storage errors persist. Never log prescription bodies or Redis/API tokens.
