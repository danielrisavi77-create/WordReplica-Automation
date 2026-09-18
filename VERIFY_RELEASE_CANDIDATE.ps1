[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WordEvidencePath,

    [Parameter(Mandatory = $true)]
    [string]$ReleaseManifestPath,

    [Parameter(Mandatory = $true)]
    [string]$ReleaseExePath,

    [string]$ExpectedSourceCommit = ''
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if ($env:OS -ne 'Windows_NT') {
    throw 'VERIFY_RELEASE_CANDIDATE.ps1 must run on Windows because it rechecks Authenticode.'
}

foreach ($path in @($WordEvidencePath, $ReleaseManifestPath, $ReleaseExePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required release evidence is missing: $path"
    }
}

if ([string]::IsNullOrWhiteSpace($ExpectedSourceCommit)) {
    $ExpectedSourceCommit = (& git -C $PSScriptRoot rev-parse HEAD).Trim().ToLowerInvariant()
}
$ExpectedSourceCommit = $ExpectedSourceCommit.Trim().ToLowerInvariant()
if ($ExpectedSourceCommit -notmatch '^[a-f0-9]{40}$') {
    throw 'ExpectedSourceCommit must be a full 40-character Git SHA.'
}

$wordEvidence = Get-Content -LiteralPath $WordEvidencePath -Raw | ConvertFrom-Json
$releaseManifest = Get-Content -LiteralPath $ReleaseManifestPath -Raw | ConvertFrom-Json

if ([string]$wordEvidence.sourceCommit -ne $ExpectedSourceCommit) {
    throw 'Word evidence source commit mismatch'
}
if ([string]$wordEvidence.wordGateOutcome -ne 'success') {
    throw "Real Microsoft Word gate did not succeed: $($wordEvidence.wordGateOutcome)"
}

if ([string]$releaseManifest.sourceCommit -ne $ExpectedSourceCommit) {
    throw 'Release manifest source commit mismatch'
}
if ($releaseManifest.sourceTreeClean -ne $true) {
    throw 'Release manifest does not attest sourceTreeClean=true.'
}
if (-not ([string]$releaseManifest.sourceBranch).StartsWith('release/')) {
    throw 'Release manifest source branch is not a release/ branch.'
}

$exeFile = Get-Item -LiteralPath $ReleaseExePath
$actualHash = (Get-FileHash -LiteralPath $ReleaseExePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ([string]$releaseManifest.sha256 -ne $actualHash) {
    throw 'Release executable SHA-256 mismatch'
}
if ([int64]$releaseManifest.sizeBytes -ne [int64]$exeFile.Length) {
    throw 'Release executable size mismatch'
}

$signature = Get-AuthenticodeSignature -FilePath $ReleaseExePath
if ($signature.Status -ne 'Valid') {
    throw "Signature status is not Valid: $($signature.Status) $($signature.StatusMessage)"
}
if ($null -eq $signature.SignerCertificate) {
    throw 'Release executable has no signer certificate.'
}
$actualThumbprint = ([string]$signature.SignerCertificate.Thumbprint).Replace(' ', '').ToUpperInvariant()
$manifestThumbprint = ([string]$releaseManifest.signingCertificateThumbprint).Replace(' ', '').ToUpperInvariant()
if ($actualThumbprint -ne $manifestThumbprint) {
    throw 'Signer thumbprint mismatch'
}

$summary = [ordered]@{
    status = 'PROMOTION_READY'
    sourceCommit = $ExpectedSourceCommit
    releaseSha256 = $actualHash
    signerThumbprint = $actualThumbprint
    wordGateOutcome = [string]$wordEvidence.wordGateOutcome
}
Write-Host 'WORD REPLICA PROMOTION READY' -ForegroundColor Green
$summary | ConvertTo-Json -Compress
