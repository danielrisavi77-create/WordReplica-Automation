from __future__ import annotations

import json
import hashlib
import time
from pathlib import Path
from types import SimpleNamespace

import pytest
from word_replica.repair_contract.signature import decode_spki
from word_replica.runner.http_transport import TransportError
from word_replica.runner.lekta_claim import LaunchTicket
from word_replica.runner.one_shot import OneShotRunner, OneShotRunnerConfig
from word_replica.runner.secure_retry import SecureRetryStore


FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"
SOURCE_BYTES = b"PK-public-repair-contract-v1"
TARGET_BYTES = b"PK-public-repair-contract-v1-target"


class FakeTransport:
    def __init__(self, contract: dict, order: list[str]) -> None:
        self.contract = contract
        self.order = order

    def post_json(self, _url: str, payload: dict) -> dict:
        if payload.get("event"):
            return {
                "ok": True,
                "jobId": payload["jobId"],
                "localState": "processing" if payload["event"] == "processing" else payload["event"],
                "sequence": payload["sequence"],
            }
        self.order.append("claim")
        return {
            "ok": True,
            "jobId": payload["jobId"],
            "contract": self.contract,
            "sourceDownloadUrl": "https://storage.example/source",
            "targetDownloadUrl": "https://storage.example/target",
            "expiresInSeconds": 300,
        }

    def get_bytes(self, url: str, *, maximum_bytes: int) -> bytes:
        assert maximum_bytes == 20 * 1024 * 1024
        if url.endswith("/source"):
            self.order.append("source")
            return SOURCE_BYTES
        self.order.append("target")
        return TARGET_BYTES


def test_preflight_precedes_claim_and_temp_secrets_are_removed(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    captured: dict[str, Path] = {}

    class FakeService:
        def run(self, request, **_kwargs):
            order.append("word")
            captured.update({
                "original": request.original_path,
                "target": request.target_path,
                "contract": request.contract_path,
                "public_key": request.public_key_path,
            })
            assert request.original_path.read_bytes() == SOURCE_BYTES
            assert request.target_path.read_bytes() == TARGET_BYTES
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished")
            output_hash = hashlib.sha256(output.read_bytes()).hexdigest()
            return SimpleNamespace(
                full_pass=True,
                status="FULL_PASS",
                output_path=output,
                output_sha256=output_hash,
                to_json=lambda: {"status": "FULL_PASS", "output_sha256": output_hash},
            )

    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=FakeTransport(contract, order),
        package_service=FakeService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=SecureRetryStore(tmp_path / "retry"),
    )
    result = runner.run(
        LaunchTicket.parse({"version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43}),
        OneShotRunnerConfig(
                claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
                status_endpoint="https://project.supabase.co/functions/v1/repair-local-status",
                output_dir=tmp_path / "out",
        ),
    )

    assert order == ["preflight", "claim", "source", "target", "word"]
    assert result.report.full_pass is True
    assert result.output_path.read_bytes() == b"finished"
    assert all(not path.exists() for path in captured.values())


def test_invalid_contract_signature_stops_before_document_download(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    contract["sourceSize"] += 1
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=FakeTransport(contract, order),
        package_service=SimpleNamespace(run=lambda *_a, **_k: None),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=SecureRetryStore(tmp_path / "retry"),
    )

    try:
        runner.run(
            LaunchTicket.parse({"version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43}),
            OneShotRunnerConfig(
                claim_endpoint="https://project.supabase.co/claim",
                status_endpoint="https://project.supabase.co/status",
                output_dir=tmp_path / "out",
            ),
        )
    except Exception:
        pass
    else:
        raise AssertionError("tampered contract must fail")

    assert order == ["preflight", "claim"]


def test_claim_retries_when_background_storage_is_not_ready(
    tmp_path: Path, monkeypatch,
) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))

    class DelayedClaimTransport(FakeTransport):
        claim_attempts = 0

        def post_json(self, url: str, payload: dict) -> dict:
            if not payload.get("event"):
                self.claim_attempts += 1
                if self.claim_attempts == 1:
                    raise TransportError("HTTPS request failed")
            return super().post_json(url, payload)

    class FakeService:
        def run(self, request, **_kwargs):
            order.append("word")
            request.output_dir.mkdir(parents=True, exist_ok=True)
            output = request.output_dir / "repaired.docx"
            output.write_bytes(b"finished")
            output_hash = hashlib.sha256(output.read_bytes()).hexdigest()
            return SimpleNamespace(
                full_pass=True,
                status="FULL_PASS",
                output_path=output,
                output_sha256=output_hash,
                to_json=lambda: {"status": "FULL_PASS", "output_sha256": output_hash},
            )

        @staticmethod
        def retry_checkpoint_sha256(_job_id: str) -> None:
            return None

    monkeypatch.setattr(time, "sleep", lambda _seconds: None)
    transport = DelayedClaimTransport(contract, order)
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=FakeService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=SecureRetryStore(tmp_path / "retry"),
    )

    result = runner.run(
        LaunchTicket.parse({"version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43}),
        OneShotRunnerConfig(
            claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
            status_endpoint="https://project.supabase.co/functions/v1/repair-local-status",
            output_dir=tmp_path / "out",
        ),
    )

    assert transport.claim_attempts == 2


def test_crash_after_server_claim_reuses_dpapi_identity_before_second_claim(
    tmp_path: Path,
) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    ticket = LaunchTicket.parse({
        "version": 1,
        "jobId": contract["jobId"],
        "claimToken": "A" * 43,
    })
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint="https://project.supabase.co/functions/v1/repair-local-status",
        output_dir=tmp_path / "out",
    )
    retry_store = SecureRetryStore(tmp_path / "retry")
    first_device: list[str] = []

    class CrashAfterAcceptedTransport(FakeTransport):
        def post_json(self, url: str, payload: dict) -> dict:
            if not payload.get("event"):
                first_device.append(payload["devicePublicKeySpki"])
                raise RuntimeError("process crashed after server accepted claim")
            return super().post_json(url, payload)

    first_runner = OneShotRunner(
        preflight=lambda: order.append("preflight-1"),
        transport=CrashAfterAcceptedTransport(contract, order),
        package_service=SimpleNamespace(run=lambda *_a, **_k: None),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    with pytest.raises(RuntimeError, match="server accepted claim"):
        first_runner.run(ticket, config)

    assert not retry_store.path_for(ticket.job_id).exists()
    assert retry_store.claim_attempt_path_for(ticket.job_id).is_file()

    second_device: list[str] = []

    class RecoveringTransport(FakeTransport):
        def post_json(self, url: str, payload: dict) -> dict:
            if not payload.get("event"):
                second_device.append(payload["devicePublicKeySpki"])
            return super().post_json(url, payload)

    class FakeService:
        def run(self, request, **_kwargs):
            order.append("word")
            request.output_dir.mkdir(parents=True, exist_ok=True)
            output = request.output_dir / "repaired.docx"
            output.write_bytes(b"finished")
            output_hash = hashlib.sha256(output.read_bytes()).hexdigest()
            return SimpleNamespace(
                full_pass=True,
                status="FULL_PASS",
                output_path=output,
                output_sha256=output_hash,
                to_json=lambda: {"status": "FULL_PASS", "output_sha256": output_hash},
            )

        @staticmethod
        def retry_checkpoint_sha256(_job_id: str) -> None:
            return None

    second_runner = OneShotRunner(
        preflight=lambda: order.append("preflight-2"),
        transport=RecoveringTransport(contract, order),
        package_service=FakeService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
        identity_factory=lambda: (_ for _ in ()).throw(AssertionError("new device generated")),
    )
    result = second_runner.run(ticket, config)

    assert first_device == second_device
    assert result.output_path.read_bytes() == b"finished"
    assert not retry_store.claim_attempt_path_for(ticket.job_id).exists()


def test_claim_binding_survives_crash_before_launch_pointer_is_sanitized(
    tmp_path: Path,
) -> None:
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    ticket = LaunchTicket.parse({
        "version": 1,
        "jobId": contract["jobId"],
        "claimToken": "A" * 43,
    })
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: None,
        transport=FakeTransport(contract, []),
        package_service=SimpleNamespace(run=lambda *_a, **_k: None),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )

    with pytest.raises(RuntimeError, match="pointer write crashed"):
        runner.run(
            ticket,
            OneShotRunnerConfig(
                claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
                status_endpoint="https://project.supabase.co/functions/v1/repair-local-status",
                output_dir=tmp_path / "out",
            ),
            on_claimed=lambda _job_id: (_ for _ in ()).throw(
                RuntimeError("pointer write crashed")
            ),
        )

    assert retry_store.path_for(ticket.job_id).is_file()
    assert retry_store.claim_attempt_path_for(ticket.job_id).is_file()
    assert retry_store.load(ticket.job_id).identity.public_key_spki == (
        retry_store.load_claim_attempt(ticket.job_id).identity.public_key_spki
    )
