from __future__ import annotations

import base64
import json
from datetime import datetime, timezone

import pytest
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import encode_dss_signature

from word_replica.runner.lekta_claim import DeviceIdentity
from word_replica.runner.lekta_status import (
    StatusProtocolError,
    build_signed_local_repair_status,
    send_local_repair_status,
    submit_signed_local_repair_status,
)


JOB_ID = "33333333-3333-4333-8333-333333333333"
STATUS_ENDPOINT = "https://project.supabase.co/functions/v1/repair-local-status"


def _decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


class FakeTransport:
    def __init__(self, response: dict) -> None:
        self.response = response
        self.calls: list[tuple[str, dict]] = []

    def post_json(self, url: str, payload: dict) -> dict:
        self.calls.append((url, payload))
        return self.response


def test_status_payload_is_canonical_p1363_signed_by_claimed_device() -> None:
    identity = DeviceIdentity.generate()
    transport = FakeTransport({
        "ok": True,
        "jobId": JOB_ID,
        "localState": "processing",
        "sequence": 1,
    })

    receipt = send_local_repair_status(
        STATUS_ENDPOINT,
        identity,
        transport,
        job_id=JOB_ID,
        sequence=1,
        event="processing",
        occurred_at=datetime(2026, 8, 25, 10, 0, tzinfo=timezone.utc),
    )

    assert receipt.local_state == "processing"
    assert len(transport.calls) == 1
    url, payload = transport.calls[0]
    assert url == STATUS_ENDPOINT
    signature = _decode(payload.pop("signature"))
    assert len(signature) == 64
    assert payload == {
        "version": 1,
        "jobId": JOB_ID,
        "sequence": 1,
        "event": "processing",
        "occurredAt": "2026-08-25T10:00:00.000Z",
        "checkpointSha256": None,
        "outputSha256": None,
        "reportSha256": None,
    }
    r = int.from_bytes(signature[:32], "big")
    s = int.from_bytes(signature[32:], "big")
    canonical = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    identity.private_key.public_key().verify(
        encode_dss_signature(r, s), canonical, ec.ECDSA(hashes.SHA256())
    )


def test_completed_status_requires_hashes_before_transport() -> None:
    transport = FakeTransport({})

    with pytest.raises(StatusProtocolError, match="invalid status hashes"):
        send_local_repair_status(
            STATUS_ENDPOINT,
            DeviceIdentity.generate(),
            transport,
            job_id=JOB_ID,
            sequence=2,
            event="completed",
            occurred_at=datetime(2026, 8, 25, 10, 0, tzinfo=timezone.utc),
        )

    assert transport.calls == []


def test_exact_signed_completion_can_be_retried_without_private_key() -> None:
    signed = build_signed_local_repair_status(
        DeviceIdentity.generate(),
        job_id=JOB_ID,
        sequence=2,
        event="completed",
        occurred_at=datetime(2026, 8, 25, 10, 1, tzinfo=timezone.utc),
        output_sha256="a" * 64,
        report_sha256="b" * 64,
    )
    original = json.loads(json.dumps(signed))
    response = {
        "ok": True,
        "jobId": JOB_ID,
        "localState": "completed",
        "sequence": 2,
    }
    transport = FakeTransport(response)

    first = submit_signed_local_repair_status(STATUS_ENDPOINT, transport, signed)
    second = submit_signed_local_repair_status(STATUS_ENDPOINT, transport, signed)

    assert first == second
    assert transport.calls == [(STATUS_ENDPOINT, original), (STATUS_ENDPOINT, original)]
    assert signed == original


@pytest.mark.parametrize("response", [
    {"ok": True, "jobId": JOB_ID, "localState": "processing", "sequence": 2},
    {"ok": True, "jobId": "44444444-4444-4444-8444-444444444444", "localState": "processing", "sequence": 1},
    {"ok": True, "jobId": JOB_ID, "localState": "completed", "sequence": 1},
])
def test_status_rejects_unbound_or_wrong_state_receipt(response: dict) -> None:
    with pytest.raises(StatusProtocolError, match="invalid status response"):
        send_local_repair_status(
            STATUS_ENDPOINT,
            DeviceIdentity.generate(),
            FakeTransport(response),
            job_id=JOB_ID,
            sequence=1,
            event="processing",
            occurred_at=datetime(2026, 8, 25, 10, 0, tzinfo=timezone.utc),
        )
