# AGENTS.md

## Purpose
This repository may be maintained with both OpenAI Codex and Claude. Treat the repository, Git history, README.md, and PROJECT_STATUS.md as the source of truth.

## Before making changes
1. Read README.md.
2. Read PROJECT_STATUS.md.
3. Read CLAUDE.md for cross-agent context.
4. Inspect `git status` and the recent commit history.
5. Inspect the relevant code before editing.

## Project baseline
- Windows desktop prescription application.
- Python desktop UI with bilingual English/Arabic support.
- PDF and editable Word export.
- Cloud QR workflow backed by the deployed `rx-v2` Next.js/Redis service.
- Current release baseline at the time this file was added: v8.0.
- Patient history/settings contain privacy-sensitive data and use Windows DPAPI protections.

## Rules
- Continue the existing implementation; do not rebuild or redesign completed features unless explicitly requested.
- Preserve existing behavior and compatibility unless the task explicitly requires a change.
- Do not weaken encryption, privacy controls, validation, QR security, or error handling.
- Do not expose, hard-code, commit, or log secrets/API keys.
- Do not use production patient data in tests.
- Preserve Arabic/RTL behavior and existing English behavior.
- Preserve the deployed cloud payload contract unless a task explicitly changes it.
- Avoid destructive migrations or storage-format changes without explicit approval.
- Prefer small, reviewable changes.
- Run the relevant automated tests before finishing.
- Update documentation when behavior changes.

## Handoff protocol
Before ending a coding session or when another AI will take over:
1. Update PROJECT_STATUS.md.
2. Record what was completed.
3. Record files changed.
4. Record tests/checks run and their results.
5. Record unresolved bugs or risks.
6. Record the exact recommended next step.
7. Commit/push the work when the environment permits.

Do not assume chat history will be available to the next coding agent.
