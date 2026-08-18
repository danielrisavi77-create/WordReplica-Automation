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

function New-WordReplicaShortcut {
    param([string]$ShortcutPath)

    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $installExe
        $shortcut.WorkingDirectory = $installDir
        $shortcut.IconLocation = "$installExe,0"
        $shortcut.Description = "Word Replica - local DOCX reconstruction and QA"
        $shortcut.Save()
    } finally {
        if ($null -ne $shortcut) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut)
        }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
    }

    if (-not (Test-Path -LiteralPath $ShortcutPath -PathType Leaf)) {
        throw "Shortcut was not created: $ShortcutPath"
    }
}

$desktop = [Environment]::GetFolderPath("Desktop")
$desktopShortcut = Join-Path $desktop "Word Replica.lnk"
New-WordReplicaShortcut -ShortcutPath $desktopShortcut

$startMenuDir = Join-Path ([Environment]::GetFolderPath("StartMenu")) "Programs"
$startMenuShortcut = Join-Path $startMenuDir "Word Replica.lnk"
New-WordReplicaShortcut -ShortcutPath $startMenuShortcut

Write-Host "WORD REPLICA INSTALL: PASS" -ForegroundColor Green
Write-Host "Installed: $installExe"
Write-Host "Desktop shortcut: $desktopShortcut"
Write-Host "Start Menu shortcut: $startMenuShortcut"
Write-Host "SHA-256: $($installedHash.ToLowerInvariant())"
