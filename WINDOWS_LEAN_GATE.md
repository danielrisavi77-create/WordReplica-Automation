# Windows Lean Release Gate

The release gate installs core and test dependencies only. PySide6 and PyInstaller are optional and are not required to verify reconstruction fidelity or Word COM compatibility.

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\RUN_WINDOWS_RELEASE_GATE.ps1
```

The gate uses `.venv_gate`, disables pip caching, checks free disk space, and stops immediately if dependency installation fails.
