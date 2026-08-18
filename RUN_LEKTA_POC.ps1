param(
    [string]$LocalRoot = "C:\WordReplica-Automation",
    [string]$PackageDir = "C:\WordReplica-Automation\poc\kalogjera\package",
    [string]$DiagnosticsRoot = "C:\WordReplica-Automation\diagnostics",
    [ValidateSet("word", "pure-docx")]
    [string]$Renderer = "word",
    [int]$StopAfterEvent
)
$ErrorActionPreference = "Stop"
$Repo = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $LocalRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "Automation venv is missing: $VenvPython."
}
Push-Location $Repo
try {
    $harnessArgs = @(
        "scripts\run_lekta_poc.py",
        "--package-dir", $PackageDir,
        "--diagnostics-root", $DiagnosticsRoot,
        "--renderer", $Renderer,
        "--repo-root", $Repo,
        "--python", $VenvPython
    )
    if ($PSBoundParameters.ContainsKey("StopAfterEvent")) {
        $harnessArgs += @("--stop-after-event", $StopAfterEvent)
    }
    # Foreground and synchronous on purpose: no background job, no spawned window.
    & $VenvPython @harnessArgs
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
