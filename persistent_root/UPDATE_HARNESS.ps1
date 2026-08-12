$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Updates = Join-Path $Root 'updates'
$Backup = Join-Path $Root 'backup'
$Current = Join-Path $Root 'current'
$VersionPath = Join-Path $Root 'version.json'
$Venv = Join-Path $Root '.venv_harness'
$Py = Join-Path $Venv 'Scripts\python.exe'
$Cli = Join-Path $Current 'scripts\persistent_harness\update_cli.py'
foreach ($Dir in @($Updates,$Backup)) {
    if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
}
foreach ($Path in @($Current,$VersionPath,$Cli)) { if (-not (Test-Path $Path)) { throw "Required persistent harness path missing: $Path" } }
$Packages = @(Get-ChildItem -LiteralPath $Updates -File -Filter '*.zip')
if ($Packages.Count -ne 1) { throw "Place exactly one update ZIP in updates. Found: $($Packages.Count)" }
if (-not (Test-Path $Py)) {
    if (Get-Command py -ErrorAction SilentlyContinue) { & py -3.12 -m venv $Venv } else { & python -m venv $Venv }
    if ($LASTEXITCODE -ne 0) { throw 'Could not create reusable .venv_harness.' }
    & $Py -m pip install -e $Current --no-cache-dir
    if ($LASTEXITCODE -ne 0) { throw 'Could not initialize harness environment.' }
}
$SelectedPackage = $Packages[0].FullName
& $Py $Cli --root $Root --zip $SelectedPackage --venv-python $Py
$Code = $LASTEXITCODE
if ($Code -eq 0) {
    $Applied = Join-Path $Updates 'applied'
    New-Item -ItemType Directory -Path $Applied -Force | Out-Null
    Move-Item -LiteralPath $SelectedPackage -Destination (Join-Path $Applied ([System.IO.Path]::GetFileName($SelectedPackage))) -Force
}
exit $Code
