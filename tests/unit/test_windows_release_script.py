from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "BUILD_WINDOWS_RELEASE.ps1"


def _text() -> str:
    return SCRIPT.read_text(encoding="utf-8")


def test_windows_release_script_exists_and_runs_verified_desktop_builder():
    assert SCRIPT.is_file()
    text = _text()
    assert "BUILD_WINDOWS_APP.ps1" in text
    assert "-SkipReleaseGate" not in text
    assert "word-replica-build-manifest.json" in text


def test_windows_release_script_refuses_dirty_or_wrong_source_branch():
    text = _text()
    assert "rev-parse --abbrev-ref HEAD" in text
    assert "release/" in text
    assert "status --porcelain" in text
    assert "source tree must be clean" in text


def test_windows_release_script_supports_certificate_store_and_artifact_signing():
    text = _text()
    assert "CertificateStore" in text
    assert "ArtifactSigning" in text
    assert "SigningCertificateThumbprint" in text
    assert "ArtifactSigningSignToolPath" in text
    assert "ArtifactSigningDlibPath" in text
    assert "ArtifactSigningMetadataPath" in text
    assert "Set-AuthenticodeSignature" in text
    assert "/dlib" in text
    assert "/dmdf" in text


def test_windows_release_script_requires_timestamp_and_reverifies_signature():
    text = _text()
    assert "TimestampServer" in text
    assert "Get-AuthenticodeSignature" in text
    assert "Signature status is not Valid" in text
    assert "SignerCertificate.Thumbprint" in text


def test_windows_release_script_emits_signed_release_manifest_bound_to_source():
    text = _text()
    assert "word-replica-release-manifest.json" in text
    for field in (
        "schemaVersion = 1",
        "fileName = $exeFile.Name",
        "version = $buildManifest.version",
        "sha256 = $signedHash",
        "sizeBytes = $exeFile.Length",
        "sourceCommit = $sourceCommit",
        "sourceBranch = $sourceBranch",
        "sourceTreeClean = $true",
        "unsignedBuildSha256 = $buildManifest.sha256",
        "signingMode = $SigningMode",
        "signingCertificateThumbprint = $signerThumbprint",
        "timestampServer = $TimestampServer",
    ):
        assert field in text


def test_windows_release_script_checks_build_manifest_source_and_bytes_before_signing():
    text = _text()
    assert "build manifest sourceCommit does not match release source" in text
    assert "build manifest SHA-256 does not match unsigned EXE" in text
    assert "build manifest size does not match unsigned EXE" in text
