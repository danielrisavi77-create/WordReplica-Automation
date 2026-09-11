from __future__ import annotations

from hashlib import sha256
import subprocess
import sys
from pathlib import Path

import pytest

from word_replica.repair_contract.signature import decode_spki
from word_replica.runner.trust_store import load_trust_keys


ROOT = Path(__file__).resolve().parents[2]
FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows release build only")
def test_release_script_prepares_the_runtime_trust_asset(tmp_path: Path) -> None:
    script = ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1"
    public_key = FIXTURE_DIR / "public-key.spki.b64url"
    destination = tmp_path / "trusted_keys.json"

    completed = subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(script),
        "-PublicKeyPath", str(public_key),
        "-KeyId", "lekta-prod-test",
        "-GeneratedTrustStorePath", str(destination),
        "-RunnerPythonPath", sys.executable,
        "-PrepareOnly",
    ], cwd=ROOT, text=True, capture_output=True, timeout=30, check=False)

    assert completed.returncode == 0, completed.stdout + completed.stderr
    spki = public_key.read_text(encoding="utf-8").strip()
    assert load_trust_keys(destination) == {"lekta-prod-test": decode_spki(spki)}
    expected = sha256(decode_spki(spki)).hexdigest()
    assert f"Public key SHA-256: {expected}" in completed.stdout



def test_release_script_requires_and_verifies_authenticode_signature() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    assert "SigningCertificateThumbprint" in script
    assert "TimestampServer" in script
    assert "Cert:\\CurrentUser\\My" in script
    assert "HasPrivateKey" in script
    assert "Set-AuthenticodeSignature" in script
    assert "Get-AuthenticodeSignature" in script
    assert "Signature status nije Valid" in script


def test_release_script_emits_signed_artifact_sha256_manifest() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    assert "Get-FileHash" in script
    assert "-Algorithm SHA256" in script
    assert "lekta-repair-runner-manifest.json" in script
    assert "sha256" in script
    assert "sizeBytes" in script
    assert "signingCertificateThumbprint" in script
    assert "ConvertTo-Json" in script


def test_release_script_emits_contract_public_key_fingerprint_in_manifest_v2() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    assert "schemaVersion = 2" in script
    assert "contractPublicKeySha256 = $contractPublicKeySha256" in script
    assert "prepared.contract_public_key_sha256" in script
    assert "schemaVersion = 1" not in script


def test_release_script_packages_repair_contract_fixer_ids() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    assert "src\\word_replica\\repair_contract\\fixer_ids.json" in script
    assert "word_replica/repair_contract" in script
    assert "$fixerIdsData" in script
    assert script.count("--add-data") >= 2


def test_release_script_boot_smoke_tests_the_frozen_runner_before_signing() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    self_test = script.index("--self-test")
    signing = script.index("Set-AuthenticodeSignature")
    assert "Start-Process" in script
    assert "-WindowStyle Hidden" in script
    assert self_test < signing


def test_release_script_attests_clean_automation_dev_source_in_manifest() -> None:
    script = (ROOT / "BUILD_LEKTA_REPAIR_RUNNER.ps1").read_text(encoding="utf-8")

    assert "rev-parse --abbrev-ref HEAD" in script
    assert "automation-dev" in script
    assert "status --porcelain" in script
    assert "rev-parse HEAD" in script
    assert "sourceCommit" in script
    assert "sourceBranch" in script
    assert "sourceTreeClean" in script
    assert "engineVersion" in script
