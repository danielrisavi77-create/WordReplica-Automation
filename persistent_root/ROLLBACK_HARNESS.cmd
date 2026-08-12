@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ROLLBACK_HARNESS.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if "%ERR%"=="0" (echo HARNESS ROLLBACK: COMPLETE) else (echo HARNESS ROLLBACK: FAIL)
pause
exit /b %ERR%
