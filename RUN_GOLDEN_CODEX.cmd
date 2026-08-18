@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0RUN_GOLDEN_CODEX.ps1" %*
exit /b %ERRORLEVEL%
