param(
    [string]$AutomationRoot = "C:\WordReplica-Automation",
    [string]$RepoUrl = "https://github.com/danielrisavi77-create/WordReplica-Automation.git",
    [string]$GoldenFileName = "Glavna verzija rektorova (grupno)(1).docx",
    [string]$PackageRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Write-Step([string]$Message) {
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Run-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
    & git @Args
    if ($LASTEXITCODE -ne 0) { throw "git failed: git $($Args -join ' ')" }
}

function Copy-Tree([string]$Source, [string]$Destination) {
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
    }
}

function Expand-Fresh([string]$Archive, [string]$Destination) {
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Expand-Archive -LiteralPath $Archive -DestinationPath $Destination -Force
}

function Resolve-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3.12 -c "import sys; assert sys.version_info >= (3,12)" 2>$null
        if ($LASTEXITCODE -eq 0) { return @{Exe="py"; Prefix=@("-3.12")} }
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -c "import sys; assert sys.version_info >= (3,12)" 2>$null
        if ($LASTEXITCODE -eq 0) { return @{Exe="python"; Prefix=@()} }
    }
    throw "Python 3.12+ is required. Install Python 3.12 and run this installer again."
}

if ($env:OS -ne "Windows_NT") { throw "This bootstrap must run on Windows." }
if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "Git for Windows is required." }

$PayloadRoot = Join-Path $PackageRoot "payload"
$BaselineZip = Join-Path $PayloadRoot "baseline_main.zip"
$AutomationZip = Join-Path $PayloadRoot "automation_final.zip"
$GoldenSeed = Join-Path $PayloadRoot $GoldenFileName
foreach ($Required in @($BaselineZip, $AutomationZip)) {
    if (-not (Test-Path -LiteralPath $Required)) { throw "Missing setup payload: $Required" }
}

$RepoRoot = Join-Path $AutomationRoot "repo"
$GoldenDir = Join-Path $AutomationRoot "golden"
$GoldenTarget = Join-Path $GoldenDir $GoldenFileName
$VenvRoot = Join-Path $AutomationRoot ".venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
foreach ($Dir in @($AutomationRoot, $GoldenDir, (Join-Path $AutomationRoot "work"), (Join-Path $AutomationRoot "diagnostics"), (Join-Path $AutomationRoot "archive"), (Join-Path $AutomationRoot "state"))) {
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
}

Write-Step "Preparing local-only Golden source"
if (-not (Test-Path -LiteralPath $GoldenTarget)) {
    if (Test-Path -LiteralPath $GoldenSeed) {
        Copy-Item -LiteralPath $GoldenSeed -Destination $GoldenTarget -Force
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = "Select WordReplica GOLDEN #1 DOCX"
        $dialog.Filter = "Word documents (*.docx)|*.docx"
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Golden DOCX selection was cancelled."
        }
        Copy-Item -LiteralPath $dialog.FileName -Destination $GoldenTarget -Force
    }
}
$GoldenHashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $GoldenTarget).Hash
[Environment]::SetEnvironmentVariable("WORD_REPLICA_AUTOMATION_ROOT", $AutomationRoot, "User")
[Environment]::SetEnvironmentVariable("WORD_REPLICA_GOLDEN_DOCX", $GoldenTarget, "User")
$env:WORD_REPLICA_AUTOMATION_ROOT = $AutomationRoot
$env:WORD_REPLICA_GOLDEN_DOCX = $GoldenTarget

Write-Step "Cloning or refreshing private WordReplica repository"
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot ".git"))) {
    if (Test-Path -LiteralPath $RepoRoot) { Remove-Item -LiteralPath $RepoRoot -Recurse -Force }
    Run-Git clone $RepoUrl $RepoRoot
}
Push-Location $RepoRoot
try {
    Run-Git fetch origin --prune
    Run-Git checkout main
    Run-Git pull --ff-only origin main

    $name = [string](& git config user.name 2>$null)
    $name = $name.Trim()
    if (-not $name) { Run-Git config user.name "WordReplica Automation Bootstrap" }
    $email = [string](& git config user.email 2>$null)
    $email = $email.Trim()
    if (-not $email) { Run-Git config user.email "wordreplica-automation@users.noreply.github.com" }

    $MainAlreadySeeded = Test-Path -LiteralPath (Join-Path $RepoRoot "src\word_replica")
    if (-not $MainAlreadySeeded) {
        Write-Step "Seeding validated WordReplica 2.0.4 baseline on main"
        $BaselineTemp = Join-Path $env:TEMP ("wordreplica-baseline-" + [guid]::NewGuid().ToString("N"))
        Expand-Fresh $BaselineZip $BaselineTemp
        Copy-Tree $BaselineTemp $RepoRoot
        Remove-Item -LiteralPath $BaselineTemp -Recurse -Force
    }

    Write-Step "Creating local Python environment"
    if (-not (Test-Path -LiteralPath $VenvPython)) {
        $py = Resolve-PythonCommand
        & $py.Exe @($py.Prefix) -m venv $VenvRoot
        if ($LASTEXITCODE -ne 0) { throw "Failed to create Python virtual environment." }
    }
    & $VenvPython -m pip install -e ".[test]"
    if ($LASTEXITCODE -ne 0) { throw "pip install -e \".[test]\" failed." }

    if (-not $MainAlreadySeeded) {
        Write-Step "Verifying stable baseline before pushing main"
        & $VenvPython -m pytest -q
        if ($LASTEXITCODE -ne 0) { throw "Baseline regression suite failed. main was not pushed." }
        Run-Git add -A
        & git diff --cached --quiet
        if ($LASTEXITCODE -ne 0) { Run-Git commit -m "chore: seed WordReplica 2.0.4 automation baseline" }
        Run-Git push origin main
    }

    $RemoteAutomation = [string](& git ls-remote --heads origin automation-dev)
    $RemoteAutomation = $RemoteAutomation.Trim()
    if ($RemoteAutomation) {
        Write-Step "Using existing automation-dev without overwriting Codex work"
        if (& git branch --list automation-dev) {
            Run-Git checkout automation-dev
        } else {
            Run-Git checkout -b automation-dev --track origin/automation-dev
        }
        Run-Git pull --ff-only origin automation-dev
    } else {
        Write-Step "Creating automation-dev and applying Codex automation layer"
        Run-Git checkout -b automation-dev main
        $AutomationTemp = Join-Path $env:TEMP ("wordreplica-automation-" + [guid]::NewGuid().ToString("N"))
        Expand-Fresh $AutomationZip $AutomationTemp
        Copy-Tree $AutomationTemp $RepoRoot
        Remove-Item -LiteralPath $AutomationTemp -Recurse -Force

        & $VenvPython -m pip install -e ".[test]"
        if ($LASTEXITCODE -ne 0) { throw "Automation dependency install failed." }
        & $VenvPython -m pytest -q
        if ($LASTEXITCODE -ne 0) { throw "Automation regression suite failed. automation-dev was not pushed." }
        Run-Git add -A
        & git diff --cached --quiet
        if ($LASTEXITCODE -ne 0) { Run-Git commit -m "feat: add local Codex Golden automation" }
        Run-Git push -u origin automation-dev
    }

    [Environment]::SetEnvironmentVariable("WORD_REPLICA_AUTOMATION_ROOT", $AutomationRoot, "User")
    [Environment]::SetEnvironmentVariable("WORD_REPLICA_GOLDEN_DOCX", $GoldenTarget, "User")
    $env:WORD_REPLICA_AUTOMATION_ROOT = $AutomationRoot
    $env:WORD_REPLICA_GOLDEN_DOCX = $GoldenTarget

    Write-Step "Final regression verification on automation-dev"
    & $VenvPython -m pytest -q
    if ($LASTEXITCODE -ne 0) { throw "Final regression suite failed." }

    $GoldenHashAfter = (Get-FileHash -Algorithm SHA256 -LiteralPath $GoldenTarget).Hash
    if ($GoldenHashBefore -ne $GoldenHashAfter) { throw "Golden source changed during setup." }

    Write-Host "`nAUTOMATION SETUP: PASS" -ForegroundColor Green
    Write-Host "Repo:   $RepoRoot"
    Write-Host "Branch: automation-dev"
    Write-Host "Golden: $GoldenTarget"
    Write-Host "Golden SHA-256: $GoldenHashAfter"
    Write-Host "`nNext: open this folder in the Codex desktop app and follow AGENTS.md."
    Write-Host "Golden command: .\RUN_GOLDEN_CODEX.ps1"
} finally {
    Pop-Location
}
