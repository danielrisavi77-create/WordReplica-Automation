$ErrorActionPreference = "Stop"
$installDir = Join-Path $env:LOCALAPPDATA "WordReplica"
$desktop = [Environment]::GetFolderPath("Desktop")
$desktopShortcut = Join-Path $desktop "Word Replica.lnk"
$startMenuDir = Join-Path ([Environment]::GetFolderPath("StartMenu")) "Programs"
$startMenuShortcut = Join-Path $startMenuDir "Word Replica.lnk"
Remove-Item -Force $desktopShortcut -ErrorAction SilentlyContinue
Remove-Item -Force $startMenuShortcut -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force $installDir -ErrorAction SilentlyContinue
Write-Host "Word Replica removed from this Windows account." -ForegroundColor Green
