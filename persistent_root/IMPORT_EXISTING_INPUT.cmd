@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0IMPORT_EXISTING_INPUT.ps1"
set "ERR=%ERRORLEVEL%"
pause
exit /b %ERR%
