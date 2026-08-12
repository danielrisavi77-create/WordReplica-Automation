param(
    [string]$SourceExe = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($SourceExe)) {
    $candidates = @(
        (Join-Path $PSScriptRoot "dist\WordReplica.exe"),
        (Join-Path (Split-Path $PSScriptRoot -Parent) "dist\WordReplica.exe"),
        (Join-Path $HOME "Downloads\dist\WordReplica.exe")
    )
    $SourceExe = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
}

if ([string]::IsNullOrWhiteSpace($SourceExe)) {
    throw "Could not find the already-built WordReplica.exe. Expected it under Downloads\dist\WordReplica.exe."
}

$installer = Join-Path $PSScriptRoot "INSTALL_WORD_REPLICA.ps1"
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
    throw "INSTALL_WORD_REPLICA.ps1 is missing."
}

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer -SourceExe $SourceExe
if ($LASTEXITCODE -ne 0) {
    throw "Word Replica installation failed."
}
