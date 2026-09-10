from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

from word_replica.repair_contract.signature import decode_spki
from word_replica.runner.trust_store import load_trust_keys


ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "BUILD_LEKTA_REPAIR_RUNNER_DEV.ps1"
FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"


def test_development_build_has_a_separate_executable_path() -> None:
    assert SCRIPT.is_file(), "development executable build path does not exist"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows development build only")
def test_development_script_prepares_the_runtime_trust_asset(tmp_path: Path) -> None:
    public_key = FIXTURE_DIR / "public-key.spki.b64url"
    destination = tmp_path / "trusted_keys.json"

    completed = subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(SCRIPT),
        "-PublicKeyPath", str(public_key),
        "-KeyId", "lekta-dev-test",
        "-GeneratedTrustStorePath", str(destination),
        "-RunnerPythonPath", sys.executable,
        "-PrepareOnly",
    ], cwd=ROOT, text=True, capture_output=True, timeout=30, check=False)

    assert completed.returncode == 0, completed.stdout + completed.stderr
    spki = public_key.read_text(encoding="utf-8").strip()
    assert load_trust_keys(destination) == {"lekta-dev-test": decode_spki(spki)}


def test_development_build_preloads_security_module_before_long_running_tools() -> None:
    script = SCRIPT.read_text(encoding="utf-8")

    security_import = script.index("Import-Module (Join-Path $PSHOME")
    security_manifest = script.index(
        "'Modules\\Microsoft.PowerShell.Security\\Microsoft.PowerShell.Security.psd1'",
        security_import,
    )
    error_action = script.index("-ErrorAction Stop", security_manifest)
    utility_import = script.index("Import-Module (Join-Path $PSHOME", security_import + 1)
    utility_manifest = script.index(
        "'Modules\\Microsoft.PowerShell.Utility\\Microsoft.PowerShell.Utility.psd1'",
        utility_import,
    )
    utility_error_action = script.index("-ErrorAction Stop", utility_manifest)
    pytest = script.index("& $RunnerPythonPath -m pytest")
    pyinstaller = script.index("& $RunnerPythonPath -m PyInstaller")
    signature_check = script.index("Get-AuthenticodeSignature")
    file_hash = script.index("Get-FileHash")

    assert security_import < security_manifest < error_action < utility_import
    assert utility_import < utility_manifest < utility_error_action < pytest
    assert pytest < pyinstaller < signature_check < file_hash


def test_development_build_packages_the_release_contract_assets() -> None:
    script = SCRIPT.read_text(encoding="utf-8")

    assert "--name LektaRepairDev" in script
    assert "scripts\\lekta_repair_runner_dev_entry.py" in script
    assert "src\\word_replica\\repair_contract\\fixer_ids.json" in script
    assert "word_replica/repair_contract" in script
    assert "word_replica/runner" in script
    assert script.count("--add-data") >= 2


def test_development_build_is_unsigned_and_emits_only_a_dev_manifest() -> None:
    script = SCRIPT.read_text(encoding="utf-8")

    assert "Set-AuthenticodeSignature" not in script
    assert "lekta-repair-runner-manifest.json" not in script
    assert "lekta-repair-runner-dev-manifest.json" in script
    assert "developmentOnly = $true" in script
    assert "Get-AuthenticodeSignature" in script
    assert "Status -eq 'Valid'" in script


def test_development_build_runs_runner_regressions_and_frozen_self_test() -> None:
    script = SCRIPT.read_text(encoding="utf-8")

    for test_path in (
        "tests/unit/test_lekta_one_shot_runner.py",
        "tests/unit/test_lekta_one_shot_status_reporting.py",
        "tests/unit/test_lekta_portable_entry.py",
        "tests/unit/test_lekta_runner_claim.py",
        "tests/unit/test_lekta_runner_http.py",
        "tests/unit/test_lekta_runner_review_regressions.py",
        "tests/unit/test_lekta_runner_status.py",
        "tests/unit/test_lekta_runner_trust_store.py",
        "tests/unit/test_lekta_secure_retry_store.py",
        "tests/unit/test_lekta_word_preflight.py",
        "tests/unit/test_repair_package_service.py",
    ):
        assert test_path in script
    assert "--self-test" in script
    assert "Start-Process" in script
    assert "-WindowStyle Hidden" in script
