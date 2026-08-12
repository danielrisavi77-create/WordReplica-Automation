param(
    [switch]$LogsOnly,
    [string]$InputDir,
    [string]$ResultsRoot,
    [string]$HarnessVersion,
    [string]$ConfigPath,
    [string]$VenvPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $InputDir) { $InputDir = Join-Path $Root 'realworld_input' }
if (-not $ResultsRoot) { $ResultsRoot = Join-Path $Root 'remote_results' }
if (-not $ConfigPath) { $ConfigPath = Join-Path $Root 'harness_config.json' }
if (-not $VenvPath) { $VenvPath = Join-Path $Root '.venv_harness' }
if (-not $HarnessVersion) { $HarnessVersion = '0.1.0' }
$Py = Join-Path $VenvPath 'Scripts\python.exe'
$ReadyMarker = Join-Path $VenvPath '.word_replica_ready'

Write-Host '=== Word Replica Remote Windows/Word Test Harness ==='

$Required = @(
    (Join-Path $Root 'pyproject.toml'),
    $ConfigPath,
    (Join-Path $Root 'scripts\remote_harness\main.py'),
    (Join-Path $Root 'src\word_replica\__init__.py'),
    $InputDir
)
foreach ($Path in $Required) {
    if (-not (Test-Path $Path)) {
        throw "Required harness file/folder is missing: $Path."
    }
}

$Docx = @(Get-ChildItem -LiteralPath $InputDir -File -Filter '*.docx' -ErrorAction SilentlyContinue)
if ($Docx.Count -lt 1) { throw 'No .docx files found in persistent realworld_input.' }
Write-Host ("Documents found: {0}" -f $Docx.Count)

if (-not (Test-Path $Py)) {
    Write-Host 'Creating reusable lean harness environment...'
    $Created = $false
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3.12 -m venv $VenvPath
        if ($LASTEXITCODE -eq 0) { $Created = $true }
    }
    if (-not $Created) {
        if (-not (Get-Command python -ErrorAction SilentlyContinue)) { throw 'Python 3.12+ was not found.' }
        & python -m venv $VenvPath
        if ($LASTEXITCODE -ne 0) { throw 'Could not create reusable .venv_harness.' }
    }
}

$env:PIP_NO_CACHE_DIR = '1'
$env:PYTHONUTF8 = '1'
if (-not (Test-Path $ReadyMarker)) {
    & $Py -m pip install -e $Root --no-cache-dir
    if ($LASTEXITCODE -ne 0) { throw 'Harness dependency installation failed.' }
    New-Item -ItemType File -Path $ReadyMarker -Force | Out-Null
}

$Args = @(
    (Join-Path $Root 'scripts\remote_harness\main.py'),
    '--input-dir', $InputDir,
    '--config', $ConfigPath,
    '--result-parent', $ResultsRoot,
    '--harness-version', $HarnessVersion
)
if ($LogsOnly) { $Args += '--logs-only' }

& $Py @Args
$Code = $LASTEXITCODE
if ($Code -ne 0) { exit $Code }
Write-Host 'REMOTE WORD HARNESS: COMPLETE'
exit 0
