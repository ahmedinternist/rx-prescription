# CLAUDE.md

## Purpose
You may receive this repository after work performed by OpenAI Codex. Continue from the current repository state rather than from assumptions about prior chat history.

## Startup checklist
1. Read README.md.
2. Read PROJECT_STATUS.md.
3. Read AGENTS.md.
4. Inspect recent Git commits and `git status`.
5. Inspect the relevant implementation before making changes.

## Project context
- Bilingual English/Arabic Windows prescription desktop application.
- Python desktop UI.
- PDF and editable Word prescription output.
- Short cloud QR links use the deployed `rx-v2` Next.js/Redis viewer/backend.
- Current release baseline when this file was added: v8.0.
- Local patient history and sensitive settings use Windows DPAPI protection.

## Working rules
- Treat the current code and repository documentation as authoritative.
- Do not restart, rewrite, or redesign completed work unless specifically asked.
- Preserve compatibility, existing workflows, Arabic/RTL behavior, and export behavior unless the requested change requires otherwise.
- Do not weaken encryption, privacy, validation, cloud-link safeguards, or secret handling.
- Never commit API keys, credentials, tokens, or production patient data.
- Do not modify persistent data formats, encryption behavior, cloud payload contracts, or migrations without explicit approval.
- Prefer focused changes with tests over broad refactors.
- Run relevant tests before finishing.

## Handoff back to Codex
Before ending the session:
1. Update PROJECT_STATUS.md.
2. Summarize completed work.
3. List files changed.
4. List tests/checks run and results.
5. List unresolved issues.
6. State the exact next recommended task.
7. Commit/push when possible.

Assume Codex will not have access to this Claude conversation.
