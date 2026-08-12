@echo off
setlocal
cd /d "%~dp0"
echo === Word Replica Build and Install ===

rem Explorer can run a .cmd directly from inside a ZIP by extracting only that file
rem to a temporary directory. Refuse that mode because the companion files are absent.
if not exist "%~dp0pyproject.toml" goto :not_extracted
if not exist "%~dp0BUILD_WINDOWS_APP.ps1" goto :not_extracted
if not exist "%~dp0INSTALL_WORD_REPLICA.ps1" goto :not_extracted

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BUILD_WINDOWS_APP.ps1"
if errorlevel 1 goto :fail
if not exist "%~dp0dist\WordReplica.exe" goto :fail

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL_WORD_REPLICA.ps1"
if errorlevel 1 goto :fail
if not exist "%LOCALAPPDATA%\WordReplica\WordReplica.exe" goto :fail

echo.
echo WORD REPLICA BUILD + INSTALL: PASS
echo Installed EXE: %LOCALAPPDATA%\WordReplica\WordReplica.exe
echo You can now launch Word Replica from your Desktop shortcut.
pause
exit /b 0

:not_extracted
echo.
echo WORD REPLICA BUILD + INSTALL: NOT STARTED
echo.
echo This launcher is being run without the rest of the project files.
echo This normally happens when it is opened directly from inside the ZIP archive.
echo.
echo Please close this window, right-click the ZIP in Downloads, choose Extract All,
echo open the extracted folder, and then double-click START_HERE.cmd.
echo.
echo No application was built or installed.
pause
exit /b 2

:fail
echo.
echo WORD REPLICA BUILD + INSTALL: FAIL
echo The application was not confirmed as installed.
pause
exit /b 1
