# Word Replica v1 — Windows Desktop App Release Package

This package adds a lightweight Tkinter desktop application and Windows EXE build/install workflow on top of the RC2 engine that passed the complete Windows + Microsoft Word release gate (97/97) on the target machine.

## New desktop files

- `src/word_replica/desktop.py` — Tkinter desktop application.
- `BUILD_WINDOWS_APP.ps1` — full-gate + PyInstaller single-file EXE builder.
- `BUILD_AND_INSTALL_WORD_REPLICA.cmd` — one-click build and per-user install.
- `INSTALL_WORD_REPLICA.ps1` — copies the EXE to `%LOCALAPPDATA%\WordReplica` and creates a Desktop shortcut.
- `UNINSTALL_WORD_REPLICA.ps1` — removes the per-user install and shortcut.
- `RUN_WORD_REPLICA_FROM_SOURCE.ps1/.cmd` — optional source-mode launcher.
- `WINDOWS_APP.md` — user/build documentation.

## Integrity boundary

The desktop app calls the existing `RebuildService`; it does not change the RC2 parser, renderer, QA, audit, metadata, save-history, or source-integrity engine. It does not simulate manual typing or fabricate timestamps, editing time, revision/save history, or provenance.

## Local verification in the packaging environment

- Desktop-specific tests: 6 passed.
- Entire non-Windows suite: 75 passed, 28 Windows/Word tests skipped, 0 failed.
- Python compile gate: passed.
- RC2 engine Python files: unchanged; the only new Python production module is `word_replica.desktop`.

The final Windows EXE itself must be built on Windows because PyInstaller does not cross-compile a Windows executable from this Linux packaging environment. `BUILD_WINDOWS_APP.ps1` therefore runs the full Windows + Word gate again before creating `dist\WordReplica.exe`.
