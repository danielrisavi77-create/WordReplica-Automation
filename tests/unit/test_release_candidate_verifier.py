from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "VERIFY_RELEASE_CANDIDATE.ps1"


def _text() -> str:
    return SCRIPT.read_text(encoding="utf-8")


def test_release_candidate_verifier_requires_exact_word_and_release_evidence():
    assert SCRIPT.is_file()
    text = _text()
    assert "WordEvidencePath" in text
    assert "ReleaseManifestPath" in text
    assert "ReleaseExePath" in text
    assert "ExpectedSourceCommit" in text
    assert "wordGateOutcome" in text
    assert "sourceCommit" in text


def test_release_candidate_verifier_fails_closed_on_source_or_hash_drift():
    text = _text()
    assert "Word evidence source commit mismatch" in text
    assert "Release manifest source commit mismatch" in text
    assert "Release executable SHA-256 mismatch" in text
    assert "sourceTreeClean" in text


def test_release_candidate_verifier_rechecks_authenticode_identity():
    text = _text()
    assert "Get-AuthenticodeSignature" in text
    assert "Signature status is not Valid" in text
    assert "signingCertificateThumbprint" in text
    assert "Signer thumbprint mismatch" in text


def test_release_candidate_verifier_reports_only_explicit_promotion_ready_terminal():
    text = _text()
    assert "WORD REPLICA PROMOTION READY" in text
    assert "PROMOTION_READY" in text
