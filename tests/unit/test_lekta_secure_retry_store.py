from __future__ import annotations

import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import pytest
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import ec

from word_replica.runner.lekta_claim import DeviceIdentity
from word_replica.runner.lekta_status import build_signed_local_repair_status
from word_replica.runner.secure_retry import (
    RetryState,
    RetryStateError,
    SecureRetryStore,
    WindowsDpapiProtector,
)


JOB_ID = "33333333-3333-4333-8333-333333333333"
SOURCE_HASH = "a" * 64
CONTRACT_HASH = "b" * 64


def _state(identity: DeviceIdentity, **overrides) -> RetryState:
    values = {
        "job_id": JOB_ID,
        "source_sha256": SOURCE_HASH,
        "contract_sha256": CONTRACT_HASH,
        "engine_version": "0.1.0",
        "contract_version": 1,
        "phase": "processing",
        "sequence": 1,
        "checkpoint_sha256": None,
        "output_path": None,
        "output_sha256": None,
        "report_sha256": None,
        "pending_status": None,
        "identity": identity,
    }
    values.update(overrides)
    return RetryState.build(**values)


@pytest.mark.skipif(sys.platform != "win32", reason="Windows DPAPI only")
def test_windows_dpapi_round_trip_is_bound_to_entropy() -> None:
    protector = WindowsDpapiProtector()
    ciphertext = protector.protect(b"private-retry-state", entropy=b"job-one")

    assert ciphertext != b"private-retry-state"
    assert protector.unprotect(ciphertext, entropy=b"job-one") == b"private-retry-state"
    with pytest.raises(RetryStateError, match="DPAPI unprotect failed"):
        protector.unprotect(ciphertext, entropy=b"job-two")


def test_store_encrypts_identity_and_binding_then_restores_same_device(tmp_path: Path) -> None:
    identity = DeviceIdentity.generate()
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())
    state = _state(identity)

    path = store.save(state)
    on_disk = path.read_bytes()

    assert SOURCE_HASH.encode("ascii") not in on_disk
    assert CONTRACT_HASH.encode("ascii") not in on_disk
    assert identity.public_key_spki.encode("ascii") not in on_disk
    loaded = store.load(JOB_ID)
    assert loaded.job_id == JOB_ID
    assert loaded.source_sha256 == SOURCE_HASH
    assert loaded.phase == "processing"
    assert loaded.sequence == 1
    assert loaded.identity.public_key_spki == identity.public_key_spki
    challenge = b"same-private-key"
    signature = loaded.identity.private_key.sign(challenge, ec.ECDSA(hashes.SHA256()))
    identity.private_key.public_key().verify(signature, challenge, ec.ECDSA(hashes.SHA256()))


def test_store_fails_closed_on_ciphertext_tampering(tmp_path: Path) -> None:
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())
    path = store.save(_state(DeviceIdentity.generate()))
    data = bytearray(path.read_bytes())
    data[-8] ^= 1
    path.write_bytes(data)

    with pytest.raises(RetryStateError, match="invalid encrypted retry state"):
        store.load(JOB_ID)


def test_store_refuses_to_replace_immutable_job_binding(tmp_path: Path) -> None:
    identity = DeviceIdentity.generate()
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())
    original = _state(identity)
    store.save(original)

    with pytest.raises(RetryStateError, match="retry binding mismatch"):
        store.save(_state(identity, source_sha256="c" * 64, sequence=2))

    assert store.load(JOB_ID).source_sha256 == SOURCE_HASH


def test_claim_attempt_first_writer_wins_and_secrets_stay_encrypted(tmp_path: Path) -> None:
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())
    first_identity = DeviceIdentity.generate()
    second_identity = DeviceIdentity.generate()
    claim_token = "A" * 43

    path = store.save_claim_attempt(
        job_id=JOB_ID, claim_token=claim_token, identity=first_identity
    )
    store.save_claim_attempt(
        job_id=JOB_ID, claim_token=claim_token, identity=second_identity
    )

    on_disk = path.read_bytes()
    assert claim_token.encode("ascii") not in on_disk
    assert first_identity.public_key_spki.encode("ascii") not in on_disk
    loaded = store.load_claim_attempt(JOB_ID)
    assert loaded.claim_token == claim_token
    assert loaded.identity.public_key_spki == first_identity.public_key_spki
    assert loaded.identity.public_key_spki != second_identity.public_key_spki


def test_delete_removes_only_exact_job_state_and_keeps_unrelated_files(tmp_path: Path) -> None:
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())
    store.save(_state(DeviceIdentity.generate()))
    workspace = store.workspace_for(JOB_ID)
    workspace.mkdir()
    (workspace / "temporary-target.docx").write_bytes(b"temporary")
    unrelated = tmp_path / "keep.txt"
    unrelated.write_text("keep", encoding="utf-8")

    store.delete(JOB_ID)

    assert not store.path_for(JOB_ID).exists()
    assert unrelated.read_text(encoding="utf-8") == "keep"


def test_pending_retryable_receipt_is_bound_to_checkpoint_and_device(tmp_path: Path) -> None:
    identity = DeviceIdentity.generate()
    checkpoint_sha256 = "e" * 64
    pending_status = build_signed_local_repair_status(
        identity,
        job_id=JOB_ID,
        sequence=2,
        event="retryable",
        occurred_at=datetime(2026, 8, 25, 10, 1, tzinfo=timezone.utc),
        checkpoint_sha256=checkpoint_sha256,
    )
    store = SecureRetryStore(tmp_path, protector=WindowsDpapiProtector())

    store.save(_state(
        identity,
        phase="retryable",
        sequence=2,
        checkpoint_sha256=checkpoint_sha256,
        pending_status=pending_status,
    ))

    loaded = store.load(JOB_ID)
    assert loaded.pending_status == pending_status
    assert loaded.checkpoint_sha256 == checkpoint_sha256


def test_pending_completion_receipt_is_encrypted_and_restored_exactly(tmp_path: Path) -> None:
    identity = DeviceIdentity.generate()
    pending_status = build_signed_local_repair_status(
        identity,
        job_id=JOB_ID,
        sequence=2,
        event="completed",
        occurred_at=datetime(2026, 8, 25, 10, 1, tzinfo=timezone.utc),
        output_sha256="c" * 64,
        report_sha256="d" * 64,
    )
    output_path = str((tmp_path / "result.docx").resolve())
    store = SecureRetryStore(tmp_path / "state", protector=WindowsDpapiProtector())
    state = _state(
        identity,
        phase="completed_pending_receipt",
        sequence=2,
        output_path=output_path,
        output_sha256="c" * 64,
        report_sha256="d" * 64,
        pending_status=pending_status,
    )

    path = store.save(state)

    on_disk = path.read_bytes()
    assert pending_status["signature"].encode("ascii") not in on_disk
    assert ("c" * 64).encode("ascii") not in on_disk
    loaded = store.load(JOB_ID)
    assert loaded.pending_status == pending_status
