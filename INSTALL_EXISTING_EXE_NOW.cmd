@echo off
setlocal
cd /d "%~dp0"
echo === Word Replica Install Existing EXE ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL_EXISTING_EXE_NOW.ps1"
if errorlevel 1 goto :fail
if not exist "%LOCALAPPDATA%\WordReplica\WordReplica.exe" goto :fail
echo.
echo WORD REPLICA INSTALL EXISTING EXE: PASS
echo Installed EXE: %LOCALAPPDATA%\WordReplica\WordReplica.exe
echo Desktop shortcut: %USERPROFILE%\Desktop\Word Replica.lnk
pause
exit /b 0
:fail
echo.
echo WORD REPLICA INSTALL EXISTING EXE: FAIL
pause
exit /b 1
