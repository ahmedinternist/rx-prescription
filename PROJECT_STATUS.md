# PROJECT_STATUS.md

## Current repository state
- Repository: `ahmedinternist/rx-prescription`
- Branch: `main`
- Current release baseline: **v8.0**
- Latest release commit when this file was created:
  - `df369a9011df80ae316d826a641d8f2a6ed5e624`
  - **Release v8.0 with stability, performance, and workflow improvements**
  - Date: 2026-10-02

## Application summary
This is a bilingual English/Arabic Windows prescription desktop application that supports:
- prescription entry and patient history,
- PDF and editable Word export,
- drug database/search and treatment templates,
- cloud QR prescription links,
- the deployed `rx-v2` Next.js/Redis viewer/backend.

Sensitive local patient/settings data uses Windows DPAPI protections. The cloud QR workflow uploads only the defined prescription payload and uses short HTTPS links.

## Current baseline notes
- README.md reflects the v8.0 executable/build naming.
- New cloud links use the deployed viewer/backend and a 60-day retention policy for records created after the v5.7 backend change.
- Unfinished prescriptions are not restored on startup.
- Existing patient history remains local and encrypted.
- Recent development has included stability, workflow, performance, cloud QR, clinic identity, and export improvements.

## Active task
No unfinished coding task is assumed in this file.

When beginning a new task, replace this section with:
- the exact objective,
- relevant files/modules,
- constraints,
- acceptance criteria.

## Last handoff
Initial Codex/Claude handoff infrastructure added.

Files added:
- `AGENTS.md`
- `CLAUDE.md`
- `PROJECT_STATUS.md`

No application code was intentionally changed.

## Tests/checks for this handoff
Documentation-only change. No application runtime behavior changed.

## Known risks / do not change casually
- Windows DPAPI encryption and persistence behavior.
- Patient-history storage and privacy behavior.
- Cloud API keys or other secrets.
- Cloud QR payload/viewer compatibility.
- Arabic/RTL rendering.
- PDF/Word output compatibility.
- Existing database/import compatibility.

## Next-step template
At the end of every future Codex or Claude session, update this file:

### Objective
<what was requested>

### Completed
- ...

### Files changed
- ...

### Tests/checks
- ...

### Unresolved issues
- ...

### Exact next step
- ...
