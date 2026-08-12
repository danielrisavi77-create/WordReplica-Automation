param([string]$SourceDir)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Current = Join-Path $Root 'current'
$Destination = Join-Path $Root 'realworld_input'
$Venv = Join-Path $Root '.venv_harness'
$Py = Join-Path $Venv 'Scripts\python.exe'
if (-not $SourceDir) { $SourceDir = Read-Host 'Path to old realworld_input folder' }
if (-not (Test-Path $Py)) {
    if (Get-Command py -ErrorAction SilentlyContinue) { & py -3.12 -m venv $Venv } else { & python -m venv $Venv }
    if ($LASTEXITCODE -ne 0) { throw 'Could not create reusable .venv_harness.' }
}
$Code = "from pathlib import Path; import sys; sys.path.insert(0,r'$Current'); from scripts.persistent_harness.migrate import import_docx_corpus; r=import_docx_corpus(Path(r'$SourceDir'),Path(r'$Destination')); print('Copied:',r.copied,'Skipped identical:',r.skipped_identical)"
& $Py -c $Code
exit $LASTEXITCODE
