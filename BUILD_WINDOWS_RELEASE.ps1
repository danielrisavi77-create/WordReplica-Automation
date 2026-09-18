[CmdletBinding()]
param(
    [ValidateSet('CertificateStore', 'ArtifactSigning')]
    [string]$SigningMode = 'CertificateStore',
    [string]$SigningCertificateThumbprint = '',
    [string]$ExpectedPublisherThumbprint = '',
    [Parameter(Mandatory = $true)]
    [string]$TimestampServer,
    [string]$ArtifactSigningSignToolPath = '',
    [string]$ArtifactSigningDlibPath = '',
    [string]$ArtifactSigningMetadataPath = ''
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if ($env:OS -ne 'Windows_NT') {
    throw 'BUILD_WINDOWS_RELEASE.ps1 must be run on Windows.'
}

$sourceBranch = (& git -C $PSScriptRoot rev-parse --abbrev-ref HEAD).Trim()
if ($LASTEXITCODE -ne 0) {
    throw 'Could not read release source branch.'
}
if (-not $sourceBranch.StartsWith('release/')) {
    throw "Production release must run from a release/ branch; got: $sourceBranch"
}

$sourceCommit = (& git -C $PSScriptRoot rev-parse HEAD).Trim().ToLowerInvariant()
if ($LASTEXITCODE -ne 0 -or $sourceCommit -notmatch '^[a-f0-9]{40}$') {
    throw 'Could not read a valid full release source commit.'
}

$sourceStatus = @(& git -C $PSScriptRoot status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0) {
    throw 'Could not read release source tree status.'
}
if ($sourceStatus.Count -ne 0) {
    throw 'Production release source tree must be clean.'
}

if ([string]::IsNullOrWhiteSpace($TimestampServer)) {
    throw 'TimestampServer is required for a production release.'
}

$expectedPublisher = $ExpectedPublisherThumbprint.Replace(' ', '').ToUpperInvariant()
$certificate = $null

if ($SigningMode -eq 'CertificateStore') {
    if ([string]::IsNullOrWhiteSpace($SigningCertificateThumbprint)) {
        throw 'SigningCertificateThumbprint is required for CertificateStore mode.'
    }
    $certificateThumbprint = $SigningCertificateThumbprint.Replace(' ', '').ToUpperInvariant()
    if ($certificateThumbprint -notmatch '^[A-F0-9]{40}$') {
        throw 'SigningCertificateThumbprint must be a 40-character SHA-1 certificate thumbprint.'
    }
    $certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$certificateThumbprint" -ErrorAction Stop
    if (-not $certificate.HasPrivateKey) {
        throw 'Signing certificate does not have a private key.'
    }
    if ([string]::IsNullOrWhiteSpace($expectedPublisher)) {
        $expectedPublisher = $certificateThumbprint
    }
} else {
    foreach ($requiredPath in @(
        $ArtifactSigningSignToolPath,
        $ArtifactSigningDlibPath,
        $ArtifactSigningMetadataPath
    )) {
        if ([string]::IsNullOrWhiteSpace($requiredPath) -or -not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw 'ArtifactSigning requires SignTool, Azure.CodeSigning.Dlib.dll, and metadata.json.'
        }
    }
    if ($expectedPublisher -notmatch '^[A-F0-9]{40}$') {
        throw 'ExpectedPublisherThumbprint is required for ArtifactSigning mode.'
    }
}

Write-Host '=== Word Replica Production Release ===' -ForegroundColor Cyan
Write-Host "Source branch: $sourceBranch"
Write-Host "Source commit: $sourceCommit"

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'BUILD_WINDOWS_APP.ps1')
if ($LASTEXITCODE -ne 0) {
    throw 'Verified desktop build failed.'
}

$exePath = Join-Path $PSScriptRoot 'dist\WordReplica.exe'
$buildManifestPath = Join-Path $PSScriptRoot 'dist\word-replica-build-manifest.json'
if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw 'Verified desktop build did not produce WordReplica.exe.'
}
if (-not (Test-Path -LiteralPath $buildManifestPath -PathType Leaf)) {
    throw 'Verified desktop build did not produce word-replica-build-manifest.json.'
}

$buildManifest = Get-Content -LiteralPath $buildManifestPath -Raw | ConvertFrom-Json
if ([string]$buildManifest.sourceCommit -ne $sourceCommit) {
    throw 'build manifest sourceCommit does not match release source'
}

$unsignedFile = Get-Item -LiteralPath $exePath
$unsignedHash = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ([string]$buildManifest.sha256 -ne $unsignedHash) {
    throw 'build manifest SHA-256 does not match unsigned EXE'
}
if ([int64]$buildManifest.sizeBytes -ne [int64]$unsignedFile.Length) {
    throw 'build manifest size does not match unsigned EXE'
}

if ($SigningMode -eq 'ArtifactSigning') {
    $signToolArguments = @(
        'sign',
        '/v',
        '/debug',
        '/fd', 'SHA256',
        '/tr', $TimestampServer,
        '/td', 'SHA256',
        '/dlib', $ArtifactSigningDlibPath,
        '/dmdf', $ArtifactSigningMetadataPath,
        $exePath
    )
    & $ArtifactSigningSignToolPath @signToolArguments
    if ($LASTEXITCODE -ne 0) {
        throw 'ArtifactSigning failed.'
    }
} else {
    $signature = Set-AuthenticodeSignature -FilePath $exePath -Certificate $certificate -HashAlgorithm SHA256 -TimestampServer $TimestampServer
    if ($signature.Status -ne 'Valid') {
        throw "Signature status is not Valid: $($signature.Status) $($signature.StatusMessage)"
    }
}

$verifiedSignature = Get-AuthenticodeSignature -FilePath $exePath
if ($verifiedSignature.Status -ne 'Valid') {
    throw "Signature status is not Valid after verification: $($verifiedSignature.Status)"
}
if ($null -eq $verifiedSignature.SignerCertificate) {
    throw 'Signed executable does not expose a signer certificate.'
}

$signerThumbprint = ([string]$verifiedSignature.SignerCertificate.Thumbprint).Replace(' ', '').ToUpperInvariant()
if ($signerThumbprint -notmatch '^[A-F0-9]{40}$') {
    throw 'SignerCertificate.Thumbprint is not a valid certificate thumbprint.'
}
if ($signerThumbprint -ne $expectedPublisher) {
    throw 'SignerCertificate.Thumbprint does not match ExpectedPublisherThumbprint.'
}

$exeFile = Get-Item -LiteralPath $exePath
$signedHash = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
$releaseManifestPath = Join-Path $PSScriptRoot 'dist\word-replica-release-manifest.json'
$temporaryManifestPath = "$releaseManifestPath.tmp"

$releaseManifest = [ordered]@{
    schemaVersion = 1
    fileName = $exeFile.Name
    version = $buildManifest.version
    sha256 = $signedHash
    sizeBytes = $exeFile.Length
    sourceCommit = $sourceCommit
    sourceBranch = $sourceBranch
    sourceTreeClean = $true
    unsignedBuildSha256 = $buildManifest.sha256
    signingMode = $SigningMode
    signingCertificateThumbprint = $signerThumbprint
    timestampServer = $TimestampServer
}

$releaseManifest | ConvertTo-Json | Set-Content -LiteralPath $temporaryManifestPath -Encoding utf8
Move-Item -LiteralPath $temporaryManifestPath -Destination $releaseManifestPath -Force

Write-Host ''
Write-Host 'WORD REPLICA PRODUCTION RELEASE: PASS' -ForegroundColor Green
Write-Host "EXE: $exePath"
Write-Host "Release SHA-256: $signedHash"
Write-Host "Signer: $signerThumbprint"
Write-Host "Manifest: $releaseManifestPath"
