# Word Replica v1 — Windows Release Candidate

This release candidate replaces the fragile all-COM reconstruction path with a hybrid strategy:

1. The deterministic canonical DOCX engine reconstructs content, styles, numbering, tables, sections, media bytes, headers/footers, notes, bookmarks and field instructions.
2. Microsoft Word is opened through COM as a real compatibility gate against the exact reconstructed package.
3. Word validation is read-only, so Word cannot normalize away retained package parts during the gate.
4. Save history and `WordReplicaActualSaveCount` remain truthful; no simulated typing, artificial delays, fake revisions, or fabricated edit-time metadata are used.

## One-command Windows gate

From PowerShell in the extracted project directory:

```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```

The script creates `.venv` if necessary, installs dependencies, enables real Word tests, compiles the code, and runs the complete pytest suite.
