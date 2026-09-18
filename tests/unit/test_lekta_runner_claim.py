from __future__ import annotations

import base64

import pytest
from cryptography.hazmat.primitives.serialization import load_der_public_key

from word_replica.runner.lekta_claim import (
    ClaimProtocolError,
    DeviceIdentity,
    LaunchTicket,
    claim_local_repair,
)


JOB_ID = "33333333-3333-4333-8333-333333333333"
CLAIM_TOKEN = "A" * 43


class FakeTransport:
    def __init__(self, response: dict) -> None:
        self.response = response
        self.calls: list[tuple[str, dict]] = []

    def post_json(self, url: str, payload: dict) -> dict:
        self.calls.append((url, payload))
        return self.response


def _b64url_decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def test_device_identity_is_a_fresh_p256_spki() -> None:
    first = DeviceIdentity.generate()
    second = DeviceIdentity.generate()

    first_public = load_der_public_key(_b64url_decode(first.public_key_spki))
    assert first_public.curve.name == "secp256r1"
    assert first.public_key_spki != second.public_key_spki
    assert first.private_key is not None


def test_claim_posts_only_bound_job_token_and_device_key() -> None:
    identity = DeviceIdentity.generate()
    ticket = LaunchTicket.parse({"version": 1, "jobId": JOB_ID, "claimToken": CLAIM_TOKEN})
    response_contract = {"contractVersion": 1, "jobId": JOB_ID}
    transport = FakeTransport({
        "ok": True,
        "jobId": JOB_ID,
        "contract": response_contract,
        "sourceDownloadUrl": "https://storage.example/source",
        "targetDownloadUrl": "https://storage.example/target",
        "expiresInSeconds": 300,
    })

    claimed = claim_local_repair(
        "https://project.supabase.co/functions/v1/repair-local-claim",
        ticket,
        identity,
        transport,
    )

    assert transport.calls == [(
        "https://project.supabase.co/functions/v1/repair-local-claim",
        {
            "jobId": JOB_ID,
            "claimToken": CLAIM_TOKEN,
            "devicePublicKeySpki": identity.public_key_spki,
        },
    )]
    assert claimed.job_id == JOB_ID
    assert claimed.contract == response_contract
    assert CLAIM_TOKEN not in repr(claimed)


@pytest.mark.parametrize("field,value", [
    ("jobId", "44444444-4444-4444-8444-444444444444"),
    ("sourceDownloadUrl", "http://storage.example/source"),
    ("targetDownloadUrl", "file:///target.docx"),
    ("expiresInSeconds", 301),
])
def test_claim_rejects_unbound_or_unsafe_server_response(field: str, value: object) -> None:
    identity = DeviceIdentity.generate()
    ticket = LaunchTicket.parse({"version": 1, "jobId": JOB_ID, "claimToken": CLAIM_TOKEN})
    response = {
        "ok": True,
        "jobId": JOB_ID,
        "contract": {"contractVersion": 1, "jobId": JOB_ID},
        "sourceDownloadUrl": "https://storage.example/source",
        "targetDownloadUrl": "https://storage.example/target",
        "expiresInSeconds": 300,
    }
    response[field] = value

    with pytest.raises(ClaimProtocolError):
        claim_local_repair("https://project.supabase.co/claim", ticket, identity, FakeTransport(response))
