@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0RUN_WORD_REPLICA_FROM_SOURCE.ps1"
if errorlevel 1 pause
