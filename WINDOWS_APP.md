# Word Replica Windows Desktop App

The desktop application is a thin Windows UI over the same `RebuildService` engine used by the CLI and the release acceptance suite.

## Build the EXE

From PowerShell in the extracted project folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\BUILD_WINDOWS_APP.ps1
```

The builder:

1. verifies at least 900 MB of free temporary disk space;
2. creates a lean `.venv_appbuild` without PySide6;
3. installs core/test/build dependencies with pip cache disabled;
4. runs the complete Windows + Microsoft Word release gate;
5. builds a single-file, windowed `WordReplica.exe` with PyInstaller;
6. prints the EXE SHA-256;
7. removes the temporary build environment and build folder by default.

The finished executable is:

```text
dist\WordReplica.exe
```

Use `-SkipReleaseGate` only if the identical source tree has already passed the Windows release gate and you intentionally want to rebuild the EXE without re-running it. Use `-KeepBuildEnvironment` if you want to preserve the temporary build environment for debugging.

## Desktop workflow

1. Click **Odaberi DOCX**.
2. Choose **Automatski**, **Microsoft Word**, or **Pure DOCX**.
3. Choose **U pozadini** or **Vidljivo** for Word runs.
4. Choose **Čista replika** or **Puna vjernost**.
5. Choose **Novi metapodaci** or **Sačuvaj legitimne**.
6. Click **Pokreni rekonstrukciju**.
7. After completion, use the buttons to open the reconstructed DOCX, QA report, or project folder.

The source stays read-only by default. The advanced overwrite checkbox requires an explicit confirmation and the engine creates and verifies a backup first.

Word Replica records real saves and audit events. It does not simulate manual typing or fabricate editing time, revision history, timestamps, or provenance.

## One-click build and install

For the simplest Windows workflow, double-click:

```text
BUILD_AND_INSTALL_WORD_REPLICA.cmd
```

It runs the verified build gate, creates `dist\WordReplica.exe`, copies the finished application to:

```text
%LOCALAPPDATA%\WordReplica\WordReplica.exe
```

and creates a **Word Replica** shortcut on the current user's Desktop. Administrator rights are not required.

To remove the installed application later:

```powershell
powershell -ExecutionPolicy Bypass -File .\UNINSTALL_WORD_REPLICA.ps1
```
