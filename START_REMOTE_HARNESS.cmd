@echo off
setlocal
cd /d "%~dp0"
echo === Word Replica Remote Word Harness ===
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0RUN_REMOTE_WORD_HARNESS.ps1"
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" (
  echo.
  echo REMOTE WORD HARNESS: FAIL
  echo Exit code: %ERR%
) else (
  echo.
  echo REMOTE WORD HARNESS: COMPLETE
  echo Upload the generated WordReplica-Remote-Results-*.zip file to ChatGPT.
)
echo.
pause
exit /b %ERR%
