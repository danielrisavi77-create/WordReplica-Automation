param(
    [string]$PythonPath = ""
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== Word Replica Windows Release Gate ===" -ForegroundColor Cyan

$ownsGateVenv = [string]::IsNullOrWhiteSpace($PythonPath)
if ($ownsGateVenv) {
    $gateVenv = Join-Path $PSScriptRoot ".venv_gate"
    $venvPython = Join-Path $gateVenv "Scripts\python.exe"

    $driveName = ([System.IO.Path]::GetPathRoot($PSScriptRoot)).TrimEnd('\').TrimEnd(':')
    $drive = Get-PSDrive -Name $driveName
    $freeMB = [math]::Floor($drive.Free / 1MB)
    Write-Host "Free space on $($drive.Name): $freeMB MB"
    if ($freeMB -lt 500) {
        Write-Host "WORD REPLICA WINDOWS RELEASE GATE: BLOCKED - less than 500 MB free." -ForegroundColor Red
        exit 28
    }

    if (Test-Path $gateVenv) {
        Remove-Item -Recurse -Force $gateVenv
    }
    python -m venv $gateVenv
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    $env:PIP_NO_CACHE_DIR = "1"
    & $venvPython -m pip install --disable-pip-version-check --no-cache-dir -e ".[test]"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "WORD REPLICA WINDOWS RELEASE GATE: INSTALL FAIL" -ForegroundColor Red
        exit $LASTEXITCODE
    }
} else {
    $venvPython = $PythonPath
    if (-not (Test-Path $venvPython)) {
        throw "PythonPath does not exist: $venvPython"
    }
}

$env:WORD_REPLICA_WORD_TESTS = "1"

Write-Host "Running compile gate..." -ForegroundColor Cyan
& $venvPython -m compileall -q src tests
if ($LASTEXITCODE -ne 0) {
    Write-Host "WORD REPLICA WINDOWS RELEASE GATE: COMPILE FAIL" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "=== INSTANT GATE ===" -ForegroundColor Cyan
& $venvPython -m pytest -q tests/unit tests/integration tests/acceptance/test_acceptance_matrix.py -k "not interactive"
if ($LASTEXITCODE -ne 0) {
    Write-Host "INSTANT RESULT: FAILED" -ForegroundColor Red
    Write-Host "WORD REPLICA WINDOWS RELEASE GATE: TEST FAIL" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "INSTANT RESULT: 0 failed" -ForegroundColor Green

Write-Host "=== INTERACTIVE GATE ===" -ForegroundColor Cyan
$interactivePaths = @(
    "tests/unit/test_interactive_config.py",
    "tests/unit/test_reconstruction_events.py",
    "tests/unit/test_blueprint_compiler.py",
    "tests/unit/test_interactive_word_executor.py",
    "tests/unit/test_interactive_control.py",
    "tests/unit/test_interactive_speed.py",
    "tests/unit/test_interactive_tables.py",
    "tests/unit/test_interactive_capabilities.py",
    "tests/unit/test_interactive_preflight.py",
    "tests/unit/test_interactive_rebuild_service.py",
    "tests/unit/test_interactive_checkpoints.py",
    "tests/unit/test_interactive_verification.py",
    "tests/unit/test_desktop_interactive.py",
    "tests/unit/test_word_render.py",
    "tests/acceptance/test_interactive_acceptance_matrix.py",
    "tests/integration/word/test_interactive_word_text.py",
    "tests/integration/word/test_interactive_resume.py",
    "tests/integration/word/test_interactive_word_structures.py"
)
$existingInteractivePaths = @($interactivePaths | Where-Object { Test-Path $_ })
& $venvPython -m pytest -q $existingInteractivePaths
if ($LASTEXITCODE -ne 0) {
    Write-Host "INTERACTIVE RESULT: FAILED" -ForegroundColor Red
    Write-Host "WORD REPLICA WINDOWS RELEASE GATE: TEST FAIL" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "INTERACTIVE RESULT: 0 failed" -ForegroundColor Green

Write-Host "WORD REPLICA WINDOWS RELEASE GATE: PASS" -ForegroundColor Green
exit 0
