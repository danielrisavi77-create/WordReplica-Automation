param(
    [string]$LocalRoot = "C:\WordReplica-Automation",
    [switch]$VisibleWord
)
$ErrorActionPreference = "Stop"
$Repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $LocalRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "Automation venv is missing: $VenvPython. Run INSTALL_CODEX_AUTOMATION.cmd first."
}
Push-Location $Repo
try {
    $args = @("-m", "scripts.codex_automation.cli", "--config", (Join-Path $Repo "codex_automation.json"), "--local-root", $LocalRoot)
    if ($VisibleWord) { $args += "--visible-word" }
    & $VenvPython @args
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
