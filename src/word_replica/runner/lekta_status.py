"""Signed lifecycle events for one claimed Lekta repair job."""
from __future__ import annotations

import base64
import json
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Literal, Protocol
from urllib.parse import urlsplit

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import decode_dss_signature

from word_replica.runner.lekta_claim import DeviceIdentity


StatusEvent = Literal["processing", "heartbeat", "retryable", "completed", "local_failed"]
LocalState = Literal["claimed", "processing", "retryable", "completed", "local_failed"]

_UUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_EVENT_STATES: dict[StatusEvent, LocalState] = {
    "processing": "processing",
    "heartbeat": "processing",
    "retryable": "retryable",
    "completed": "completed",
    "local_failed": "local_failed",
}
_RESPONSE_KEYS = {"ok", "jobId", "localState", "sequence"}
_SIGNED_STATUS_KEYS = {
    "version", "jobId", "sequence", "event", "occurredAt",
    "checkpointSha256", "outputSha256", "reportSha256", "signature",
}
_TIMESTAMP = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")


class StatusProtocolError(ValueError):
    """A lifecycle event or server receipt is malformed or unbound."""


class StatusTransport(Protocol):
    def post_json(self, url: str, payload: dict) -> dict: ...


@dataclass(frozen=True, slots=True)
class LocalRepairStatusReceipt:
    job_id: str
    local_state: LocalState
    sequence: int


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _https_url(value: Any) -> str:
    if not isinstance(value, str):
        raise StatusProtocolError("status endpoint must use HTTPS")
    parsed = urlsplit(value)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise StatusProtocolError("status endpoint must use HTTPS")
    return value


def _timestamp(value: datetime) -> str:
    if not isinstance(value, datetime) or value.tzinfo is None or value.utcoffset() != timezone.utc.utcoffset(value):
        raise StatusProtocolError("invalid status timestamp")
    return value.astimezone(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def _valid_hash(value: str | None) -> bool:
    return value is None or bool(_SHA256.fullmatch(value))


def _validate_hashes(
    event: StatusEvent,
    checkpoint_sha256: str | None,
    output_sha256: str | None,
    report_sha256: str | None,
) -> None:
    if not all(_valid_hash(value) for value in (checkpoint_sha256, output_sha256, report_sha256)):
        raise StatusProtocolError("invalid status hashes")
    valid = False
    if event == "processing":
        valid = checkpoint_sha256 is None and output_sha256 is None and report_sha256 is None
    elif event == "heartbeat":
        valid = output_sha256 is None and report_sha256 is None
    elif event == "retryable":
        valid = checkpoint_sha256 is not None and output_sha256 is None and report_sha256 is None
    elif event == "completed":
        valid = output_sha256 is not None and report_sha256 is not None
    elif event == "local_failed":
        valid = output_sha256 is None and report_sha256 is None
    if not valid:
        raise StatusProtocolError("invalid status hashes")


def _canonical_bytes(payload: dict) -> bytes:
    return json.dumps(
        payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True
    ).encode("utf-8")


def _sign_p1363(identity: DeviceIdentity, payload: dict) -> str:
    der = identity.private_key.sign(_canonical_bytes(payload), ec.ECDSA(hashes.SHA256()))
    r, s = decode_dss_signature(der)
    try:
        signature = r.to_bytes(32, "big") + s.to_bytes(32, "big")
    except OverflowError as exc:
        raise StatusProtocolError("invalid device signature") from exc
    return _base64url(signature)


def build_signed_local_repair_status(
    identity: DeviceIdentity,
    *,
    job_id: str,
    sequence: int,
    event: StatusEvent,
    occurred_at: datetime,
    checkpoint_sha256: str | None = None,
    output_sha256: str | None = None,
    report_sha256: str | None = None,
) -> dict:
    if not isinstance(job_id, str) or not _UUID.fullmatch(job_id):
        raise StatusProtocolError("invalid status job")
    if isinstance(sequence, bool) or not isinstance(sequence, int) or sequence < 1:
        raise StatusProtocolError("invalid status sequence")
    if event not in _EVENT_STATES:
        raise StatusProtocolError("invalid status event")
    _validate_hashes(event, checkpoint_sha256, output_sha256, report_sha256)

    payload = {
        "version": 1,
        "jobId": job_id,
        "sequence": sequence,
        "event": event,
        "occurredAt": _timestamp(occurred_at),
        "checkpointSha256": checkpoint_sha256,
        "outputSha256": output_sha256,
        "reportSha256": report_sha256,
    }
    return {**payload, "signature": _sign_p1363(identity, payload)}


def validate_signed_local_repair_status(payload: Any) -> tuple[str, int, StatusEvent]:
    if not isinstance(payload, dict) or set(payload) != _SIGNED_STATUS_KEYS:
        raise StatusProtocolError("invalid signed status")
    job_id = payload.get("jobId")
    sequence = payload.get("sequence")
    event = payload.get("event")
    if payload.get("version") != 1 or not isinstance(job_id, str) or not _UUID.fullmatch(job_id):
        raise StatusProtocolError("invalid signed status")
    if isinstance(sequence, bool) or not isinstance(sequence, int) or sequence < 1:
        raise StatusProtocolError("invalid signed status")
    if event not in _EVENT_STATES:
        raise StatusProtocolError("invalid signed status")
    occurred_at = payload.get("occurredAt")
    if not isinstance(occurred_at, str) or not _TIMESTAMP.fullmatch(occurred_at):
        raise StatusProtocolError("invalid signed status")
    _validate_hashes(
        event,
        payload.get("checkpointSha256"),
        payload.get("outputSha256"),
        payload.get("reportSha256"),
    )
    signature = payload.get("signature")
    if not isinstance(signature, str) or not re.fullmatch(r"[A-Za-z0-9_-]{86}", signature):
        raise StatusProtocolError("invalid signed status")
    try:
        decoded = base64.urlsafe_b64decode(signature + "==")
    except (ValueError, TypeError) as exc:
        raise StatusProtocolError("invalid signed status") from exc
    if len(decoded) != 64 or _base64url(decoded) != signature:
        raise StatusProtocolError("invalid signed status")
    return job_id, sequence, event


def submit_signed_local_repair_status(
    endpoint: str,
    transport: StatusTransport,
    payload: dict,
) -> LocalRepairStatusReceipt:
    endpoint = _https_url(endpoint)
    job_id, sequence, event = validate_signed_local_repair_status(payload)
    response = transport.post_json(endpoint, dict(payload))
    expected_state = _EVENT_STATES[event]
    if (
        not isinstance(response, dict)
        or set(response) != _RESPONSE_KEYS
        or response.get("ok") is not True
        or response.get("jobId") != job_id
        or response.get("sequence") != sequence
        or response.get("localState") != expected_state
    ):
        raise StatusProtocolError("invalid status response")
    return LocalRepairStatusReceipt(
        job_id=job_id,
        local_state=expected_state,
        sequence=sequence,
    )


def send_local_repair_status(
    endpoint: str,
    identity: DeviceIdentity,
    transport: StatusTransport,
    *,
    job_id: str,
    sequence: int,
    event: StatusEvent,
    occurred_at: datetime,
    checkpoint_sha256: str | None = None,
    output_sha256: str | None = None,
    report_sha256: str | None = None,
) -> LocalRepairStatusReceipt:
    payload = build_signed_local_repair_status(
        identity,
        job_id=job_id,
        sequence=sequence,
        event=event,
        occurred_at=occurred_at,
        checkpoint_sha256=checkpoint_sha256,
        output_sha256=output_sha256,
        report_sha256=report_sha256,
    )
    return submit_signed_local_repair_status(endpoint, transport, payload)
