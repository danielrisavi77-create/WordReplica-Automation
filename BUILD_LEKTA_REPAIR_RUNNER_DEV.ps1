[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PublicKeyPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z0-9._-]{1,80}$')]
    [string]$KeyId,

    [Parameter(Mandatory = $true)]
    [string]$GeneratedTrustStorePath,

    [string]$RunnerPythonPath = '',
    [string]$OutputDirectory = '',
    [switch]$PrepareOnly
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1') -ErrorAction Stop
Import-Module (Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Utility\Microsoft.PowerShell.Utility.psd1') -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($RunnerPythonPath)) {
    $automationRoot = Split-Path -Parent $PSScriptRoot
    $RunnerPythonPath = Join-Path $automationRoot '.venv\Scripts\python.exe'
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $PSScriptRoot 'dist\lekta-runner-dev'
}

if (-not (Test-Path -LiteralPath $RunnerPythonPath -PathType Leaf)) {
    throw 'Runner Python nije dostupan.'
}

$resolvedPublicKey = (Resolve-Path -LiteralPath $PublicKeyPath -ErrorAction Stop).Path
$resolvedTrustStore = [IO.Path]::GetFullPath($GeneratedTrustStorePath)
$trustStoreParent = [IO.Path]::GetDirectoryName($resolvedTrustStore)
[void](New-Item -ItemType Directory -Force -Path $trustStoreParent)

$prepareCode = @'
from pathlib import Path
import sys
from word_replica.runner.trust_store import prepare_release_trust_store

prepare_release_trust_store(
    public_key_path=Path(sys.argv[1]),
    key_id=sys.argv[2],
    destination=Path(sys.argv[3]),
)
'@

& $RunnerPythonPath -c $prepareCode $resolvedPublicKey $KeyId $resolvedTrustStore
if ($LASTEXITCODE -ne 0) {
    throw 'Priprema javnog runner trust storea nije uspjela.'
}

if ($PrepareOnly) {
    Write-Host "Prepared: $resolvedTrustStore"
    exit 0
}

& $RunnerPythonPath -m pytest -q `
    tests/unit/test_lekta_one_shot_runner.py `
    tests/unit/test_lekta_one_shot_status_reporting.py `
    tests/unit/test_lekta_portable_entry.py `
    tests/unit/test_lekta_runner_claim.py `
    tests/unit/test_lekta_runner_http.py `
    tests/unit/test_lekta_runner_review_regressions.py `
    tests/unit/test_lekta_runner_status.py `
    tests/unit/test_lekta_runner_trust_store.py `
    tests/unit/test_lekta_secure_retry_store.py `
    tests/unit/test_lekta_word_preflight.py `
    tests/unit/test_repair_package_service.py `
    tests/unit/test_lekta_runner_dev_build.py `
    tests/unit/test_lekta_runner_dev_entry.py
if ($LASTEXITCODE -ne 0) {
    throw 'Lekta development runner regresijski gate nije prosao.'
}

$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
$buildRoot = Join-Path $PSScriptRoot 'build\lekta-runner-dev'
$trustData = "$resolvedTrustStore;word_replica/runner"
$fixerIdsPath = Join-Path $PSScriptRoot 'src\word_replica\repair_contract\fixer_ids.json'
if (-not (Test-Path -LiteralPath $fixerIdsPath -PathType Leaf)) {
    throw 'fixer_ids.json nije dostupan za development build.'
}
$fixerIdsData = "$fixerIdsPath;word_replica/repair_contract"

& $RunnerPythonPath -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name LektaRepairDev `
    --paths (Join-Path $PSScriptRoot 'src') `
    --add-data $trustData `
    --add-data $fixerIdsData `
    --distpath $resolvedOutput `
    --workpath (Join-Path $buildRoot 'work') `
    --specpath (Join-Path $buildRoot 'spec') `
    (Join-Path $PSScriptRoot 'scripts\lekta_repair_runner_dev_entry.py')
if ($LASTEXITCODE -ne 0) {
    throw 'Lekta development runner PyInstaller build nije uspio.'
}

$runnerPath = Join-Path $resolvedOutput 'LektaRepairDev.exe'
if (-not (Test-Path -LiteralPath $runnerPath -PathType Leaf)) {
    throw 'LektaRepairDev.exe nije proizveden.'
}

$selfTest = Start-Process -FilePath $runnerPath `
    -ArgumentList @('--self-test', $KeyId) -Wait -PassThru -WindowStyle Hidden
if ($selfTest.ExitCode -ne 0) {
    throw 'LektaRepairDev.exe nije prosao frozen contract self-test.'
}

$signature = Get-AuthenticodeSignature -FilePath $runnerPath
if ($signature.Status -eq 'Valid') {
    throw 'Development runner neocekivano ima valjan Authenticode potpis.'
}

$runnerFile = Get-Item -LiteralPath $runnerPath
$artifactHash = (Get-FileHash -LiteralPath $runnerPath -Algorithm SHA256).Hash.ToLowerInvariant()
$manifestPath = Join-Path $resolvedOutput 'lekta-repair-runner-dev-manifest.json'
$temporaryManifestPath = "$manifestPath.tmp"
$manifest = [ordered]@{
    schemaVersion = 1
    developmentOnly = $true
    fileName = $runnerFile.Name
    sha256 = $artifactHash
    sizeBytes = $runnerFile.Length
    contractKeyId = $KeyId
    authenticodeStatus = [string]$signature.Status
}
$manifest | ConvertTo-Json | Set-Content -LiteralPath $temporaryManifestPath -Encoding utf8
Move-Item -LiteralPath $temporaryManifestPath -Destination $manifestPath -Force

Write-Host "Built unsigned development runner: $runnerPath"
Write-Host "Development manifest: $manifestPath"
