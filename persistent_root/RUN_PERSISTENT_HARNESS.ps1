param([switch]$LogsOnly)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Current = Join-Path $Root 'current'
$InputDir = Join-Path $Root 'realworld_input'
$ResultsRoot = Join-Path $Root 'results'
$Backup = Join-Path $Root 'backup'
$Updates = Join-Path $Root 'updates'
$ConfigPath = Join-Path $Root 'harness_config.json'
$VersionPath = Join-Path $Root 'version.json'
$VenvPath = Join-Path $Root '.venv_harness'
$Runner = Join-Path $Current 'RUN_REMOTE_WORD_HARNESS.ps1'
foreach ($Dir in @($InputDir,$ResultsRoot,$Backup,$Updates)) {
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
}
foreach ($Path in @($Current,$ConfigPath,$VersionPath,$Runner)) {
    if (-not (Test-Path $Path)) { throw "Persistent harness path missing: $Path" }
}
$Version = (Get-Content -LiteralPath $VersionPath -Raw | ConvertFrom-Json).harness_version
$Args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$Runner,'-InputDir',$InputDir,'-ResultsRoot',$ResultsRoot,'-HarnessVersion',$Version,'-ConfigPath',$ConfigPath,'-VenvPath',$VenvPath)
if ($LogsOnly) { $Args += '-LogsOnly' }
& powershell.exe @Args
exit $LASTEXITCODE
