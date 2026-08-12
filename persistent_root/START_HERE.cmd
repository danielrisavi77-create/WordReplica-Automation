@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0RUN_PERSISTENT_HARNESS.ps1"
set "ERR=%ERRORLEVEL%"
echo.
if "%ERR%"=="0" (
  echo PERSISTENT REMOTE HARNESS: COMPLETE
) else (
  echo PERSISTENT REMOTE HARNESS: FAIL
  echo Exit code: %ERR%
)
pause
exit /b %ERR%
