@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0pyproject.toml" goto :not_extracted
if not exist "%~dp0BUILD_AND_INSTALL_WORD_REPLICA.cmd" goto :not_extracted
if not exist "%~dp0BUILD_WINDOWS_APP.ps1" goto :not_extracted
if not exist "%~dp0INSTALL_WORD_REPLICA.ps1" goto :not_extracted

call "%~dp0BUILD_AND_INSTALL_WORD_REPLICA.cmd"
exit /b %errorlevel%

:not_extracted
echo.
echo WORD REPLICA - EXTRACT ZIP FIRST
echo.
echo You opened START_HERE.cmd directly from the ZIP preview.
echo Windows only extracted this launcher to a temporary folder, so the project files are missing.
echo.
echo 1. Close this window.
echo 2. In Downloads, right-click WordReplica-v1-Interactive-RC1-Windows.zip.
echo 3. Choose Extract All.
echo 4. Open the extracted folder.
echo 5. Double-click START_HERE.cmd again.
echo.
echo No application was built or installed.
pause
exit /b 2
