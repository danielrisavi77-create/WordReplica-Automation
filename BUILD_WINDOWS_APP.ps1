param(
    [switch]$SkipReleaseGate,
    [switch]$KeepBuildEnvironment
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== Word Replica Windows App Builder ===" -ForegroundColor Cyan
if ($env:OS -ne "Windows_NT") {
    throw "BUILD_WINDOWS_APP.ps1 must be run on Windows."
}

$driveName = ([System.IO.Path]::GetPathRoot($PSScriptRoot)).TrimEnd('\').TrimEnd(':')
$drive = Get-PSDrive -Name $driveName
$freeMB = [math]::Floor($drive.Free / 1MB)
Write-Host "Free space on $($drive.Name): $freeMB MB"
if ($freeMB -lt 900) {
    throw "At least 900 MB free space is required for the temporary EXE build. Remove old WordReplica .venv folders and retry."
}

$venv = Join-Path $PSScriptRoot ".venv_appbuild"
$python = Join-Path $venv "Scripts\python.exe"

if (Test-Path $venv) {
    Write-Host "Removing previous app build environment..."
    Remove-Item -Recurse -Force $venv
}

Write-Host "Creating lean build environment..."
python -m venv $venv
if ($LASTEXITCODE -ne 0) { throw "Could not create build virtual environment." }

$env:PIP_NO_CACHE_DIR = "1"
& $python -m pip install --disable-pip-version-check --no-cache-dir -e ".[test,build]"
if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }

if (-not $SkipReleaseGate) {
    # Belt-and-suspenders guard: the release gate also sets this variable, but the
    # builder itself explicitly requires real Microsoft Word tests before packaging.
    $env:WORD_REPLICA_WORD_TESTS = "1"
    Write-Host "Running the complete dual Windows + Microsoft Word release gate..." -ForegroundColor Yellow
    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "RUN_WINDOWS_RELEASE_GATE.ps1") -PythonPath $python
    if ($LASTEXITCODE -ne 0) {
        throw "Release gate failed. EXE was not built."
    }
}

Write-Host "Building single-file Windows application..." -ForegroundColor Yellow
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name WordReplica `
    --icon assets\icon.ico `
    --paths src `
    --hidden-import word_replica.desktop `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --collect-submodules win32com `
    scripts\word_replica_desktop_entry.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }

$exe = Join-Path $PSScriptRoot "dist\WordReplica.exe"
if (-not (Test-Path $exe)) { throw "Build ended without dist\WordReplica.exe" }

$exeFile = Get-Item $exe
$hash = (Get-FileHash $exe -Algorithm SHA256).Hash.ToLowerInvariant()
$sizeMB = [math]::Round($exeFile.Length / 1MB, 1)
$version = (& $python -c "from word_replica import __version__; print(__version__)").Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($version)) {
    throw "Could not read Word Replica package version."
}

$sourceCommit = $null
if (Get-Command git -ErrorAction SilentlyContinue) {
    $candidateCommit = (& git -C $PSScriptRoot rev-parse HEAD 2>$null).Trim()
    if ($LASTEXITCODE -eq 0 -and $candidateCommit -match '^[a-fA-F0-9]{40}
if (-not $KeepBuildEnvironment) {
    Write-Host "Cleaning temporary build files to save disk space..."
    Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $venv -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "You can now launch: dist\WordReplica.exe" -ForegroundColor Green
) {
        $sourceCommit = $candidateCommit.ToLowerInvariant()
    }
}

$manifestPath = Join-Path $PSScriptRoot "dist\word-replica-build-manifest.json"
$manifest = [ordered]@{
    schemaVersion = 1
    fileName = $exeFile.Name
    version = $version
    sha256 = $hash
    sizeBytes = $exeFile.Length
    sourceCommit = $sourceCommit
}
$manifest | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding utf8

Write-Host ""
Write-Host "WORD REPLICA WINDOWS APP BUILD: PASS" -ForegroundColor Green
Write-Host "EXE: $exe"
Write-Host "Size: $sizeMB MB"
Write-Host "SHA-256: $hash"
Write-Host "Manifest: $manifestPath"

if (-not $KeepBuildEnvironment) {
    Write-Host "Cleaning temporary build files to save disk space..."
    Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $venv -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "You can now launch: dist\WordReplica.exe" -ForegroundColor Green
