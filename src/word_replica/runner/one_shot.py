"""One-shot orchestration from a consumed Lekta claim to WordReplica.

Preflight happens before the bearer token is consumed. The signed contract is
verified against a runner-owned trust map before either DOCX is downloaded.
The device key and lifecycle receipt stay DPAPI-protected. Source material and
the verified package stay in one runner-owned per-job workspace only while a
durable Word checkpoint is resumable; terminal success removes that workspace. Only
the new verified output is written to the user-selected output directory.
"""
from __future__ import annotations

import base64
import hashlib
import json
import re
import threading
import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Protocol

from word_replica import __version__
from word_replica.repair_contract.package import RepairPackageRequest
from word_replica.repair_contract.signature import verify_signed_contract
from word_replica.runner.http_transport import TransportError
from word_replica.runner.lekta_claim import (
    ClaimTransport,
    DeviceIdentity,
    LaunchTicket,
    claim_local_repair,
)
from word_replica.runner.lekta_status import (
    build_signed_local_repair_status,
    submit_signed_local_repair_status,
)
from word_replica.runner.secure_retry import RetryState, SecureRetryStore


MAX_DOCX_BYTES = 20 * 1024 * 1024
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_CLAIM_RETRY_DELAYS = (0.5, 1.0, 2.0, 4.0)


class OneShotTransport(ClaimTransport, Protocol):
    def get_bytes(self, url: str, *, maximum_bytes: int) -> bytes: ...


class RepairPackageServiceLike(Protocol):
    def run(self, request: RepairPackageRequest, **kwargs): ...
    def resume(self, request: RepairPackageRequest, job_id: str, **kwargs): ...
    def retry_checkpoint_sha256(self, job_id: str) -> str | None: ...


@dataclass(frozen=True, slots=True)
class OneShotRunnerConfig:
    claim_endpoint: str
    status_endpoint: str
    output_dir: Path
    heartbeat_interval_seconds: float = 120.0


@dataclass(frozen=True, slots=True)
class OneShotRunnerResult:
    report: Any
    output_path: Path


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _contract_key_id(raw: dict) -> str:
    envelope = raw.get("contractSignature")
    key_id = envelope.get("keyId") if isinstance(envelope, dict) else None
    if not isinstance(key_id, str):
        raise ValueError("signed contract has no key id")
    return key_id


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _report_sha256(report: Any) -> str:
    serializer = getattr(report, "to_json", None)
    if not callable(serializer):
        raise RuntimeError("local repair report is not serializable")
    try:
        encoded = json.dumps(
            serializer(), ensure_ascii=False, separators=(",", ":"),
            sort_keys=True, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise RuntimeError("local repair report is not serializable") from exc
    return hashlib.sha256(encoded).hexdigest()


class OneShotRunner:
    def __init__(
        self,
        *,
        preflight: Callable[[], None],
        transport: OneShotTransport,
        package_service: RepairPackageServiceLike,
        trust_keys: Mapping[str, bytes],
        retry_store: SecureRetryStore,
        identity_factory: Callable[[], DeviceIdentity] = DeviceIdentity.generate,
    ) -> None:
        self.preflight = preflight
        self.transport = transport
        self.package_service = package_service
        self.trust_keys = dict(trust_keys)
        self.retry_store = retry_store
        self.identity_factory = identity_factory

    def has_interrupted_job(self, job_id: str) -> bool:
        return self.retry_store.path_for(job_id).is_file()

    def _claim_with_retry(
        self, endpoint: str, ticket: LaunchTicket, identity: DeviceIdentity
    ):
        for attempt in range(len(_CLAIM_RETRY_DELAYS) + 1):
            try:
                return claim_local_repair(endpoint, ticket, identity, self.transport)
            except TransportError:
                if attempt == len(_CLAIM_RETRY_DELAYS):
                    raise
                time.sleep(_CLAIM_RETRY_DELAYS[attempt])
        raise AssertionError("unreachable claim retry state")

    @staticmethod
    def _state_from(
        state: RetryState,
        *,
        phase: str,
        sequence: int,
        checkpoint_sha256: str | None = None,
        output_path: str | None = None,
        output_sha256: str | None = None,
        report_sha256: str | None = None,
        pending_status: dict | None = None,
    ) -> RetryState:
        return RetryState.build(
            job_id=state.job_id,
            source_sha256=state.source_sha256,
            contract_sha256=state.contract_sha256,
            engine_version=state.engine_version,
            contract_version=state.contract_version,
            phase=phase,
            sequence=sequence,
            checkpoint_sha256=checkpoint_sha256,
            output_path=output_path,
            output_sha256=output_sha256,
            report_sha256=report_sha256,
            pending_status=pending_status,
            identity=state.identity,
        )

    def _advance_status(
        self,
        state: RetryState,
        config: OneShotRunnerConfig,
        *,
        phase: str,
        event: str,
        checkpoint_sha256: str | None = None,
    ) -> RetryState:
        sequence = state.sequence + 1
        pending = build_signed_local_repair_status(
            state.identity,
            job_id=state.job_id,
            sequence=sequence,
            event=event,
            occurred_at=datetime.now(timezone.utc),
            checkpoint_sha256=checkpoint_sha256,
        )
        pending_state = self._state_from(
            state,
            phase=phase,
            sequence=sequence,
            checkpoint_sha256=checkpoint_sha256,
            pending_status=pending,
        )
        self.retry_store.save(pending_state)
        submit_signed_local_repair_status(config.status_endpoint, self.transport, pending)
        settled = self._state_from(
            pending_state,
            phase=phase,
            sequence=sequence,
            checkpoint_sha256=checkpoint_sha256,
        )
        self.retry_store.save(settled)
        return settled

    def _prepare_claimed_request(
        self, state: RetryState, config: OneShotRunnerConfig
    ) -> RepairPackageRequest:
        if not isinstance(state.claim_contract_json, str):
            raise RuntimeError("local repair claimed package is unavailable")
        try:
            contract_bytes = state.claim_contract_json.encode("utf-8")
            raw_contract = json.loads(state.claim_contract_json)
        except (UnicodeError, ValueError) as exc:
            raise RuntimeError("local repair claimed package is unavailable") from exc
        if hashlib.sha256(contract_bytes).hexdigest() != state.contract_sha256:
            raise RuntimeError("local repair claimed contract changed")
        contract = verify_signed_contract(raw_contract, self.trust_keys)
        if contract.job_id != state.job_id or contract.source_sha256 != state.source_sha256:
            raise RuntimeError("local repair claimed binding changed")
        key_id = _contract_key_id(raw_contract)
        trusted_key = self.trust_keys[key_id]
        if state.source_download_url is None or state.target_download_url is None:
            raise RuntimeError("local repair claimed downloads are unavailable")

        workspace = self.retry_store.workspace_for(state.job_id)
        workspace.mkdir(parents=True, exist_ok=True)
        original_path = workspace / contract.source_file_name
        target_path = workspace / contract.target_file_name
        contract_path = workspace / "repair-contract.json"
        public_key_path = workspace / "lekta-public-key.spki.b64url"
        source_bytes = self.transport.get_bytes(
            state.source_download_url, maximum_bytes=MAX_DOCX_BYTES
        )
        target_bytes = self.transport.get_bytes(
            state.target_download_url, maximum_bytes=MAX_DOCX_BYTES
        )
        original_path.write_bytes(source_bytes)
        target_path.write_bytes(target_bytes)
        contract_path.write_bytes(contract_bytes)
        public_key_path.write_text(_base64url(trusted_key), encoding="ascii")
        return RepairPackageRequest(
            original_path=original_path,
            target_path=target_path,
            contract_path=contract_path,
            public_key_path=public_key_path,
            output_dir=Path(config.output_dir).resolve(),
        )

    def _workspace_request(
        self, state: RetryState, config: OneShotRunnerConfig
    ) -> RepairPackageRequest:
        workspace = self.retry_store.workspace_for(state.job_id)
        contract_path = workspace / "repair-contract.json"
        public_key_path = workspace / "lekta-public-key.spki.b64url"
        try:
            contract_bytes = contract_path.read_bytes()
            raw_contract = json.loads(contract_bytes.decode("utf-8"))
        except (OSError, UnicodeError, ValueError) as exc:
            raise RuntimeError("local repair retry package is unavailable") from exc
        if hashlib.sha256(contract_bytes).hexdigest() != state.contract_sha256:
            raise RuntimeError("local repair retry contract changed")
        contract = verify_signed_contract(raw_contract, self.trust_keys)
        if contract.job_id != state.job_id or contract.source_sha256 != state.source_sha256:
            raise RuntimeError("local repair retry binding changed")
        request = RepairPackageRequest(
            original_path=workspace / contract.source_file_name,
            target_path=workspace / contract.target_file_name,
            contract_path=contract_path,
            public_key_path=public_key_path,
            output_dir=Path(config.output_dir).resolve(),
        )
        if not request.original_path.is_file() or not request.target_path.is_file():
            raise RuntimeError("local repair retry documents are unavailable")
        return request

    def _finish_terminal_failure(
        self,
        state: RetryState,
        config: OneShotRunnerConfig,
        message: str,
        cause: Exception | None = None,
    ) -> None:
        sequence = state.sequence + 1
        pending_status = build_signed_local_repair_status(
            state.identity,
            job_id=state.job_id,
            sequence=sequence,
            event="local_failed",
            occurred_at=datetime.now(timezone.utc),
        )
        pending_state = self._state_from(
            state,
            phase="failed_pending_receipt",
            sequence=sequence,
            pending_status=pending_status,
        )
        self.retry_store.save(pending_state)
        submit_signed_local_repair_status(
            config.status_endpoint, self.transport, pending_status
        )
        self.retry_store.delete(state.job_id)
        failure = RuntimeError(message)
        if cause is None:
            raise failure
        raise failure from cause

    def _execute_package(
        self,
        operation: Callable[[], Any],
        state: RetryState,
        config: OneShotRunnerConfig,
    ) -> tuple[Any, RetryState]:
        interval = config.heartbeat_interval_seconds
        if isinstance(interval, bool) or not isinstance(interval, (int, float)) or interval <= 0:
            raise ValueError("heartbeat interval must be positive")
        stop = threading.Event()

        def send_heartbeats() -> None:
            heartbeat_state = state
            while not stop.wait(float(interval)):
                try:
                    heartbeat_state = self._advance_status(
                        heartbeat_state, config,
                        phase="processing", event="heartbeat",
                    )
                except Exception:
                    return

        thread = threading.Thread(
            target=send_heartbeats,
            name=f"lekta-heartbeat-{state.job_id}",
            daemon=True,
        )
        failure: Exception | None = None
        report: Any = None
        thread.start()
        try:
            report = operation()
        except Exception as exc:
            failure = exc
        finally:
            stop.set()
            thread.join()
        latest_state = self.retry_store.load(state.job_id)
        if failure is not None:
            try:
                checkpoint_sha256 = self.package_service.retry_checkpoint_sha256(state.job_id)
            except Exception:
                checkpoint_sha256 = None
            if isinstance(checkpoint_sha256, str) and _SHA256.fullmatch(checkpoint_sha256):
                self._advance_status(
                    latest_state, config, phase="retryable", event="retryable",
                    checkpoint_sha256=checkpoint_sha256,
                )
                raise RuntimeError("local repair is retryable from its durable checkpoint") from failure
            self._finish_terminal_failure(
                latest_state, config,
                "local repair failed without a resumable checkpoint", failure,
            )
        return report, latest_state

    def _finish_or_checkpoint(
        self,
        report: Any,
        state: RetryState,
        config: OneShotRunnerConfig,
    ) -> OneShotRunnerResult:
        if not getattr(report, "full_pass", False) or not getattr(report, "output_path", None):
            checkpoint_sha256 = self.package_service.retry_checkpoint_sha256(state.job_id)
            if not isinstance(checkpoint_sha256, str) or not _SHA256.fullmatch(checkpoint_sha256):
                self._finish_terminal_failure(
                    state, config,
                    "local repair failed without a resumable checkpoint",
                )
            self._advance_status(
                state,
                config,
                phase="retryable",
                event="retryable",
                checkpoint_sha256=checkpoint_sha256,
            )
            raise RuntimeError("local repair is retryable from its durable checkpoint")

        output_dir = Path(config.output_dir).resolve()
        output_path = Path(report.output_path).resolve()
        if not output_path.is_file() or output_path.parent != output_dir:
            raise RuntimeError("local repair output is outside the selected directory")
        output_sha256 = _sha256_file(output_path)
        reported_output_sha256 = getattr(report, "output_sha256", None)
        if reported_output_sha256 is not None and reported_output_sha256 != output_sha256:
            raise RuntimeError("local repair output hash mismatch")
        report_sha256 = _report_sha256(report)
        sequence = state.sequence + 1
        pending_status = build_signed_local_repair_status(
            state.identity,
            job_id=state.job_id,
            sequence=sequence,
            event="completed",
            occurred_at=datetime.now(timezone.utc),
            output_sha256=output_sha256,
            report_sha256=report_sha256,
        )
        pending_state = self._state_from(
            state,
            phase="completed_pending_receipt",
            sequence=sequence,
            output_path=str(output_path),
            output_sha256=output_sha256,
            report_sha256=report_sha256,
            pending_status=pending_status,
        )
        self.retry_store.save(pending_state)
        submit_signed_local_repair_status(
            config.status_endpoint, self.transport, pending_status
        )
        self.retry_store.delete(state.job_id)
        return OneShotRunnerResult(report=report, output_path=output_path)

    def run(
        self,
        ticket: LaunchTicket,
        config: OneShotRunnerConfig,
        *,
        on_claimed: Callable[[str], None] | None = None,
    ) -> OneShotRunnerResult:
        # Do not consume the one-time token until this machine has proved it can
        # create and save a benign Word document.
        self.preflight()

        claim_attempt_path = self.retry_store.claim_attempt_path_for(ticket.job_id)
        if not claim_attempt_path.is_file():
            self.retry_store.save_claim_attempt(
                job_id=ticket.job_id,
                claim_token=ticket.claim_token,
                identity=self.identity_factory(),
            )
        # Reload after publication so concurrent launches always use the
        # identity of the process that atomically won the job binding.
        claim_attempt = self.retry_store.load_claim_attempt(ticket.job_id)
        if claim_attempt.claim_token != ticket.claim_token:
            raise RuntimeError("claim attempt token changed")
        identity = claim_attempt.identity
        durable_ticket = LaunchTicket(
            job_id=ticket.job_id, claim_token=claim_attempt.claim_token
        )
        claimed = self._claim_with_retry(config.claim_endpoint, durable_ticket, identity)
        contract = verify_signed_contract(claimed.contract, self.trust_keys)
        if contract.job_id != ticket.job_id:
            raise ValueError("verified contract job mismatch")
        contract_bytes = json.dumps(
            claimed.contract, ensure_ascii=False, separators=(",", ":")
        ).encode("utf-8")
        contract_sha256 = hashlib.sha256(contract_bytes).hexdigest()
        self.retry_store.save(RetryState.build(
            job_id=ticket.job_id,
            source_sha256=contract.source_sha256,
            contract_sha256=contract_sha256,
            engine_version=__version__,
            contract_version=1,
            phase="claimed",
            sequence=0,
            checkpoint_sha256=None,
            output_path=None,
            output_sha256=None,
            report_sha256=None,
            pending_status=None,
            claim_contract_json=contract_bytes.decode("utf-8"),
            source_download_url=claimed.source_download_url,
            target_download_url=claimed.target_download_url,
            identity=identity,
        ))
        if on_claimed is not None:
            on_claimed(ticket.job_id)

        state = self.retry_store.load(ticket.job_id)
        request = self._prepare_claimed_request(state, config)
        state = self._advance_status(
            state, config, phase="processing", event="processing"
        )
        report, state = self._execute_package(
            lambda: self.package_service.run(request), state, config
        )
        return self._finish_or_checkpoint(report, state, config)

    def resume_interrupted(
        self, job_id: str, config: OneShotRunnerConfig
    ) -> OneShotRunnerResult:
        state = self.retry_store.load(job_id)
        if state.phase == "completed_pending_receipt":
            return self.resume_pending_completion(job_id, config)
        if state.phase == "failed_pending_receipt":
            return self.resume_pending_terminal_failure(job_id, config)
        if state.phase == "claimed":
            request = self._prepare_claimed_request(state, config)
            state = self._advance_status(
                state, config, phase="processing", event="processing"
            )
            report, state = self._execute_package(
                lambda: self.package_service.run(request), state, config
            )
            return self._finish_or_checkpoint(report, state, config)
        if state.pending_status is not None:
            submit_signed_local_repair_status(
                config.status_endpoint, self.transport, state.pending_status
            )
            state = self._state_from(
                state,
                phase=state.phase,
                sequence=state.sequence,
                checkpoint_sha256=state.checkpoint_sha256,
            )
            self.retry_store.save(state)

        request = self._workspace_request(state, config)
        checkpoint_sha256 = self.package_service.retry_checkpoint_sha256(job_id)
        if state.phase == "retryable":
            if checkpoint_sha256 != state.checkpoint_sha256:
                raise RuntimeError("local repair checkpoint changed")
            state = self._advance_status(
                state, config, phase="processing", event="processing"
            )
            report, state = self._execute_package(
                lambda: self.package_service.resume(request, job_id), state, config
            )
        elif state.phase == "processing":
            if checkpoint_sha256 is None:
                report, state = self._execute_package(
                    lambda: self.package_service.run(request), state, config
                )
            elif _SHA256.fullmatch(checkpoint_sha256):
                report, state = self._execute_package(
                    lambda: self.package_service.resume(request, job_id), state, config
                )
            else:
                raise RuntimeError("local repair checkpoint is invalid")
        else:
            raise RuntimeError("local repair has no resumable execution")
        return self._finish_or_checkpoint(report, state, config)

    def resume_pending_terminal_failure(
        self, job_id: str, config: OneShotRunnerConfig
    ) -> None:
        state = self.retry_store.load(job_id)
        if state.phase != "failed_pending_receipt" or state.pending_status is None:
            raise RuntimeError("local repair is not awaiting a terminal failure receipt")
        submit_signed_local_repair_status(
            config.status_endpoint, self.transport, state.pending_status
        )
        self.retry_store.delete(job_id)
        raise RuntimeError("local repair ended with a terminal local failure")

    def resume_pending_completion(
        self, job_id: str, config: OneShotRunnerConfig
    ) -> OneShotRunnerResult:
        state = self.retry_store.load(job_id)
        if state.phase != "completed_pending_receipt" or state.pending_status is None:
            raise RuntimeError("local repair is not awaiting a completion receipt")
        output_path = Path(state.output_path or "").resolve()
        output_dir = Path(config.output_dir).resolve()
        if (
            not output_path.is_file()
            or output_path.parent != output_dir
            or _sha256_file(output_path) != state.output_sha256
        ):
            raise RuntimeError("pending local repair output is unavailable or changed")
        submit_signed_local_repair_status(
            config.status_endpoint, self.transport, state.pending_status
        )
        self.retry_store.delete(job_id)
        return OneShotRunnerResult(report=None, output_path=output_path)
