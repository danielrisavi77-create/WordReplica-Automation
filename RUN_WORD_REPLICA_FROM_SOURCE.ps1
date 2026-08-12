$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$venv = Join-Path $PSScriptRoot ".venv_app"
$python = Join-Path $venv "Scripts\python.exe"
if (-not (Test-Path $python)) {
    python -m venv $venv
    $env:PIP_NO_CACHE_DIR = "1"
    & $python -m pip install --upgrade pip
    & $python -m pip install --no-cache-dir -e .
}
& $python -m word_replica.desktop
