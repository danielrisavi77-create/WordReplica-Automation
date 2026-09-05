from __future__ import annotations

import hashlib
import json
import threading
from pathlib import Path

import pytest

from word_replica.repair_contract.signature import decode_spki
from word_replica.runner.lekta_claim import LaunchTicket
from word_replica.runner.one_shot import OneShotRunner, OneShotRunnerConfig
from word_replica.runner.secure_retry import SecureRetryStore


FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"
SOURCE_BYTES = b"PK-public-repair-contract-v1"
TARGET_BYTES = b"PK-public-repair-contract-v1-target"
STATUS_ENDPOINT = "https://project.supabase.co/functions/v1/repair-local-status"


class StatusAwareTransport:
    def __init__(
        self, contract: dict, order: list[str], *,
        fail_completed_once: bool = False,
        fail_processing_once: bool = False,
        fail_local_failed_once: bool = False,
    ) -> None:
        self.contract = contract
        self.order = order
        self.statuses: list[dict] = []
        self.fail_completed_once = fail_completed_once
        self.fail_processing_once = fail_processing_once
        self.fail_local_failed_once = fail_local_failed_once

    def post_json(self, url: str, payload: dict) -> dict:
        if url == STATUS_ENDPOINT:
            self.statuses.append(payload)
            self.order.append(f"status:{payload['event']}")
            if payload["event"] == "processing" and self.fail_processing_once:
                self.fail_processing_once = False
                raise OSError("response lost after processing commit")
            if payload["event"] == "completed" and self.fail_completed_once:
                self.fail_completed_once = False
                raise OSError("offline after local FULL_PASS")
            if payload["event"] == "local_failed" and self.fail_local_failed_once:
                self.fail_local_failed_once = False
                raise OSError("response lost after local_failed commit")
            local_state = "processing" if payload["event"] == "processing" else payload["event"]
            return {
                "ok": True,
                "jobId": payload["jobId"],
                "localState": local_state,
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


class FakeReport:
    def __init__(self, output_path: Path | None, *, full_pass: bool = True) -> None:
        self.full_pass = full_pass
        self.status = "FULL_PASS" if full_pass else "RETRYABLE"
        self.output_path = str(output_path) if output_path is not None else None
        self.output_sha256 = (
            hashlib.sha256(output_path.read_bytes()).hexdigest()
            if output_path is not None else None
        )

    def to_json(self) -> dict:
        return {
            "schema_version": 1,
            "status": self.status,
            "output_sha256": self.output_sha256,
        }


def test_one_shot_reports_processing_and_verified_completion(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order)
    report_json: dict = {}

    class FakeService:
        def run(self, request, **_kwargs):
            order.append("word")
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished")
            report = FakeReport(output)
            report_json.update(report.to_json())
            return report

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
            status_endpoint=STATUS_ENDPOINT,
            output_dir=tmp_path / "out",
        ),
    )

    assert result.report.full_pass is True
    assert order == [
        "preflight", "claim", "source", "target",
        "status:processing", "word", "status:completed",
    ]
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"), (2, "completed"),
    ]
    completed = transport.statuses[1]
    assert completed["outputSha256"] == hashlib.sha256(b"finished").hexdigest()
    expected_report_hash = hashlib.sha256(
        json.dumps(report_json, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    ).hexdigest()
    assert completed["reportSha256"] == expected_report_hash
    assert not runner.retry_store.path_for(contract["jobId"]).exists()


def test_offline_completion_retries_exact_receipt_without_running_word_again(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order, fail_completed_once=True)
    word_calls = 0

    class FakeService:
        def run(self, request, **_kwargs):
            nonlocal word_calls
            word_calls += 1
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished-offline")
            return FakeReport(output)

    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=FakeService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    try:
        runner.run(ticket, config)
    except OSError as exc:
        assert "offline" in str(exc)
    else:
        raise AssertionError("first completion receipt must simulate an offline failure")

    pending = retry_store.load(contract["jobId"])
    assert pending.phase == "completed_pending_receipt"
    assert pending.pending_status == transport.statuses[-1]
    retried = runner.resume_pending_completion(contract["jobId"], config)

    assert word_calls == 1
    assert retried.output_path.read_bytes() == b"finished-offline"
    assert transport.statuses[-1] == transport.statuses[-2]
    assert not retry_store.path_for(contract["jobId"]).exists()


def test_checkpointed_failure_resumes_without_reclaiming_or_restarting_word(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order)
    checkpoint_bytes = b"durable-word-checkpoint"
    checkpoint_hash = hashlib.sha256(checkpoint_bytes).hexdigest()

    class RetryableService:
        def __init__(self) -> None:
            self.run_calls = 0
            self.resume_calls = 0
            self.first_target_path: Path | None = None

        def run(self, request, **_kwargs):
            self.run_calls += 1
            self.first_target_path = request.target_path
            order.append("word:first")
            return FakeReport(None, full_pass=False)

        def retry_checkpoint_sha256(self, job_id: str) -> str | None:
            assert job_id == contract["jobId"]
            return checkpoint_hash

        def resume(self, request, job_id: str, **_kwargs):
            self.resume_calls += 1
            assert job_id == contract["jobId"]
            assert request.target_path == self.first_target_path
            assert request.target_path.read_bytes() == TARGET_BYTES
            order.append("word:resume")
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"resumed-finished")
            return FakeReport(output)

    service = RetryableService()
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=service,
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    try:
        runner.run(ticket, config)
    except RuntimeError as exc:
        assert "retryable" in str(exc).lower()
    else:
        raise AssertionError("checkpointed first pass must request a resume")

    pending = retry_store.load(contract["jobId"])
    assert pending.phase == "retryable"
    assert pending.checkpoint_sha256 == checkpoint_hash
    assert service.first_target_path is not None and service.first_target_path.is_file()

    resumed = runner.resume_interrupted(contract["jobId"], config)

    assert resumed.output_path.read_bytes() == b"resumed-finished"
    assert service.run_calls == 1
    assert service.resume_calls == 1
    assert order.count("claim") == 1
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (2, "retryable"),
        (3, "processing"),
        (4, "completed"),
    ]
    assert not retry_store.path_for(contract["jobId"]).exists()


def test_lost_processing_response_replays_exact_status_without_a_second_claim(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order, fail_processing_once=True)

    class Service:
        def __init__(self) -> None:
            self.run_calls = 0

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

        def run(self, request, **_kwargs):
            self.run_calls += 1
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished-after-status-replay")
            return FakeReport(output)

    service = Service()
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=service,
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    with pytest.raises(OSError, match="response lost"):
        runner.run(ticket, config)

    pending = retry_store.load(contract["jobId"])
    assert pending.phase == "processing"
    assert pending.pending_status == transport.statuses[0]

    result = runner.resume_interrupted(contract["jobId"], config)

    assert result.output_path.read_bytes() == b"finished-after-status-replay"
    assert service.run_calls == 1
    assert order.count("claim") == 1
    assert transport.statuses[1] == transport.statuses[0]
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (1, "processing"),
        (2, "completed"),
    ]


def test_terminal_failure_reports_local_failed_and_removes_sensitive_workspace(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order)

    class TerminalFailureService:
        def run(self, request, **_kwargs):
            order.append("word")
            return FakeReport(None, full_pass=False)

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=TerminalFailureService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    with pytest.raises(RuntimeError, match="failed without a resumable checkpoint"):
        runner.run(ticket, config)

    assert order == [
        "preflight", "claim", "source", "target",
        "status:processing", "word", "status:local_failed",
    ]
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (2, "local_failed"),
    ]
    assert not retry_store.path_for(contract["jobId"]).exists()
    assert not retry_store.workspace_for(contract["jobId"]).exists()


def test_lost_local_failed_response_replays_exact_status_without_rerunning_word(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order, fail_local_failed_once=True)

    class TerminalFailureService:
        def __init__(self) -> None:
            self.run_calls = 0

        def run(self, request, **_kwargs):
            self.run_calls += 1
            order.append("word")
            return FakeReport(None, full_pass=False)

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

    service = TerminalFailureService()
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=service,
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    with pytest.raises(OSError, match="response lost"):
        runner.run(ticket, config)

    pending = retry_store.load(contract["jobId"])
    assert pending.phase == "failed_pending_receipt"
    assert pending.pending_status == transport.statuses[-1]

    with pytest.raises(RuntimeError, match="terminal local failure"):
        runner.resume_interrupted(contract["jobId"], config)

    assert service.run_calls == 1
    assert order.count("claim") == 1
    assert transport.statuses[-1] == transport.statuses[-2]
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (2, "local_failed"),
        (2, "local_failed"),
    ]
    assert not retry_store.path_for(contract["jobId"]).exists()
    assert not retry_store.workspace_for(contract["jobId"]).exists()


def test_unexpected_word_exception_reports_local_failed_and_preserves_cause(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order)

    class ExplodingService:
        def run(self, request, **_kwargs):
            order.append("word")
            raise OSError("Word COM disconnected")

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=ExplodingService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    with pytest.raises(RuntimeError, match="failed without a resumable checkpoint") as failure:
        runner.run(ticket, config)

    assert isinstance(failure.value.__cause__, OSError)
    assert str(failure.value.__cause__) == "Word COM disconnected"
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (2, "local_failed"),
    ]
    assert not retry_store.path_for(contract["jobId"]).exists()
    assert not retry_store.workspace_for(contract["jobId"]).exists()


def test_long_word_run_sends_signed_heartbeats_before_completion(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    heartbeat_seen = threading.Event()

    class HeartbeatTransport(StatusAwareTransport):
        def post_json(self, url: str, payload: dict) -> dict:
            response = super().post_json(url, payload)
            if url == STATUS_ENDPOINT and payload["event"] == "heartbeat":
                heartbeat_seen.set()
            return response

    class SlowService:
        def run(self, request, **_kwargs):
            order.append("word")
            assert heartbeat_seen.wait(timeout=2.0), "runner did not report progress while Word was active"
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished-after-heartbeat")
            return FakeReport(output)

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

    transport = HeartbeatTransport(contract, order)
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=SlowService(),
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
        heartbeat_interval_seconds=0.01,
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    result = runner.run(ticket, config)

    assert result.output_path.read_bytes() == b"finished-after-heartbeat"
    events = [(item["sequence"], item["event"]) for item in transport.statuses]
    assert events[0] == (1, "processing")
    assert events[-1][1] == "completed"
    assert any(event == "heartbeat" for _sequence, event in events[1:-1])
    assert [sequence for sequence, _event in events] == list(range(1, len(events) + 1))
    assert not retry_store.path_for(contract["jobId"]).exists()


def test_crash_after_claim_resumes_download_without_a_second_claim(tmp_path: Path) -> None:
    order: list[str] = []
    contract = json.loads((FIXTURE_DIR / "valid-contract.json").read_text(encoding="utf-8"))
    public_der = decode_spki((FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8"))
    transport = StatusAwareTransport(contract, order)

    class Service:
        def __init__(self) -> None:
            self.run_calls = 0

        def retry_checkpoint_sha256(self, job_id: str) -> None:
            assert job_id == contract["jobId"]
            return None

        def run(self, request, **_kwargs):
            self.run_calls += 1
            assert request.original_path.read_bytes() == SOURCE_BYTES
            assert request.target_path.read_bytes() == TARGET_BYTES
            output = request.output_dir / contract["outputPolicy"]["suggestedFileName"]
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(b"finished-after-claimed-crash")
            return FakeReport(output)

    service = Service()
    retry_store = SecureRetryStore(tmp_path / "retry")
    runner = OneShotRunner(
        preflight=lambda: order.append("preflight"),
        transport=transport,
        package_service=service,
        trust_keys={contract["contractSignature"]["keyId"]: public_der},
        retry_store=retry_store,
    )
    config = OneShotRunnerConfig(
        claim_endpoint="https://project.supabase.co/functions/v1/repair-local-claim",
        status_endpoint=STATUS_ENDPOINT,
        output_dir=tmp_path / "out",
    )
    ticket = LaunchTicket.parse({
        "version": 1, "jobId": contract["jobId"], "claimToken": "A" * 43,
    })

    def simulate_crash(_job_id: str) -> None:
        raise OSError("process stopped after durable claim")

    with pytest.raises(OSError, match="after durable claim"):
        runner.run(ticket, config, on_claimed=simulate_crash)

    claimed = retry_store.load(contract["jobId"])
    assert claimed.phase == "claimed"
    assert claimed.claim_contract_json is not None
    assert claimed.source_download_url == "https://storage.example/source"
    assert claimed.target_download_url == "https://storage.example/target"
    encrypted_state = retry_store.path_for(contract["jobId"]).read_bytes()
    assert b"https://storage.example/source" not in encrypted_state
    assert b"https://storage.example/target" not in encrypted_state
    assert b"contractSignature" not in encrypted_state
    assert not retry_store.workspace_for(contract["jobId"]).exists()

    result = runner.resume_interrupted(contract["jobId"], config)

    assert result.output_path.read_bytes() == b"finished-after-claimed-crash"
    assert service.run_calls == 1
    assert order.count("claim") == 1
    assert [(item["sequence"], item["event"]) for item in transport.statuses] == [
        (1, "processing"),
        (2, "completed"),
    ]
    assert not retry_store.path_for(contract["jobId"]).exists()
