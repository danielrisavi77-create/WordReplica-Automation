$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Current = Join-Path $Root 'current'
$Backup = Join-Path $Root 'backup'
$VersionPath = Join-Path $Root 'version.json'
$Venv = Join-Path $Root '.venv_harness'
$Py = Join-Path $Venv 'Scripts\python.exe'
$Cli = Join-Path $Current 'scripts\persistent_harness\rollback_cli.py'
foreach ($Path in @($Current,$Backup,$VersionPath,$Py,$Cli)) { if (-not (Test-Path $Path)) { throw "Required rollback path missing: $Path" } }
& $Py $Cli --root $Root
exit $LASTEXITCODE
