@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0UPDATE_HARNESS.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if "%ERR%"=="0" (echo HARNESS UPDATE: COMPLETE) else (echo HARNESS UPDATE: FAIL)
pause
exit /b %ERR%
