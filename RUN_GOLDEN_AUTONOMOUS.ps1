param(
    [string]$LocalRoot = "C:\WordReplica-Automation",
    [switch]$ResumeAfterReview
)
$ErrorActionPreference = "Stop"
$Repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $LocalRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $VenvPython)) {
    throw "Automation venv is missing: $VenvPython"
}
Push-Location $Repo
try {
    $args = @("-m", "scripts.codex_automation.autonomous_loop", "--config", (Join-Path $Repo "codex_automation.json"), "--local-root", $LocalRoot)
    if ($ResumeAfterReview) { $args += "--resume-after-review" }
    & $VenvPython @args
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
