$ErrorActionPreference = "Stop"
$installDir = Join-Path $env:LOCALAPPDATA "WordReplica"
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Word Replica.lnk"
Remove-Item -Force $shortcutPath -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force $installDir -ErrorAction SilentlyContinue
Write-Host "Word Replica removed from this Windows account." -ForegroundColor Green
