param(
    [string]$SourceExe = ""
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($SourceExe)) {
    $SourceExe = Join-Path $PSScriptRoot "dist\WordReplica.exe"
}

$sourceExe = [System.IO.Path]::GetFullPath($SourceExe)
if (-not (Test-Path -LiteralPath $sourceExe -PathType Leaf)) {
    throw "WordReplica.exe does not exist at: $sourceExe"
}

$installDir = Join-Path $env:LOCALAPPDATA "WordReplica"
$installExe = Join-Path $installDir "WordReplica.exe"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null

$sourceHash = (Get-FileHash $sourceExe -Algorithm SHA256).Hash
Copy-Item -LiteralPath $sourceExe -Destination $installExe -Force
$installedHash = (Get-FileHash $installExe -Algorithm SHA256).Hash
if ($sourceHash -ne $installedHash) {
    throw "Installed EXE hash does not match the source EXE hash."
}

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Word Replica.lnk"
$shell = New-Object -ComObject WScript.Shell
try {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $installExe
    $shortcut.WorkingDirectory = $installDir
    $shortcut.Description = "Word Replica - local DOCX reconstruction and QA"
    $shortcut.Save()
} finally {
    if ($null -ne $shortcut) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut)
    }
    if ($null -ne $shell) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
    }
}

if (-not (Test-Path -LiteralPath $shortcutPath -PathType Leaf)) {
    throw "Desktop shortcut was not created: $shortcutPath"
}

Write-Host "WORD REPLICA INSTALL: PASS" -ForegroundColor Green
Write-Host "Installed: $installExe"
Write-Host "Desktop shortcut: $shortcutPath"
Write-Host "SHA-256: $($installedHash.ToLowerInvariant())"
