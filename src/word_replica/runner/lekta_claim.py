"""Fail-closed claim protocol for one paid Lekta repair job.

This module has no Word or filesystem side effects. It creates the ephemeral
device identity and validates the complete claim response before any document
is downloaded or Microsoft Word is opened.
"""
from __future__ import annotations

import base64
import re
from dataclasses import dataclass, field
from typing import Any, Protocol
from urllib.parse import urlsplit

from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.serialization import Encoding, PublicFormat


_UUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)
_CLAIM_TOKEN = re.compile(r"^[A-Za-z0-9_-]{43}$")
_RESPONSE_KEYS = {
    "ok", "jobId", "contract", "sourceDownloadUrl", "targetDownloadUrl", "expiresInSeconds"
}


class ClaimProtocolError(ValueError):
    """The launch ticket or server response is malformed or unsafe."""


class ClaimTransport(Protocol):
    def post_json(self, url: str, payload: dict) -> dict: ...


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _https_url(value: Any) -> str:
    if not isinstance(value, str):
        raise ClaimProtocolError("signed download URL is missing")
    parsed = urlsplit(value)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise ClaimProtocolError("signed download URL must use HTTPS")
    return value


@dataclass(frozen=True, slots=True)
class LaunchTicket:
    job_id: str
    claim_token: str = field(repr=False)

    @classmethod
    def parse(cls, value: Any) -> "LaunchTicket":
        if not isinstance(value, dict) or set(value) != {"version", "jobId", "claimToken"}:
            raise ClaimProtocolError("invalid launch ticket")
        if value.get("version") != 1 or not isinstance(value.get("jobId"), str) or not _UUID.fullmatch(value["jobId"]):
            raise ClaimProtocolError("invalid launch ticket job")
        token = value.get("claimToken")
        if not isinstance(token, str) or not _CLAIM_TOKEN.fullmatch(token):
            raise ClaimProtocolError("invalid launch ticket token")
        return cls(job_id=value["jobId"], claim_token=token)


@dataclass(frozen=True, slots=True)
class DeviceIdentity:
    private_key: ec.EllipticCurvePrivateKey = field(repr=False)
    public_key_spki: str

    @classmethod
    def generate(cls) -> "DeviceIdentity":
        private_key = ec.generate_private_key(ec.SECP256R1())
        spki = private_key.public_key().public_bytes(Encoding.DER, PublicFormat.SubjectPublicKeyInfo)
        return cls(private_key=private_key, public_key_spki=_base64url(spki))


@dataclass(frozen=True, slots=True)
class ClaimedRepair:
    job_id: str
    expires_in_seconds: int
    contract: dict = field(repr=False)
    source_download_url: str = field(repr=False)
    target_download_url: str = field(repr=False)

    @classmethod
    def parse(cls, value: Any, *, expected_job_id: str) -> "ClaimedRepair":
        if not isinstance(value, dict) or set(value) != _RESPONSE_KEYS or value.get("ok") is not True:
            raise ClaimProtocolError("invalid claim response")
        job_id = value.get("jobId")
        if job_id != expected_job_id:
            raise ClaimProtocolError("claim response job mismatch")
        contract = value.get("contract")
        if not isinstance(contract, dict) or contract.get("jobId") != expected_job_id:
            raise ClaimProtocolError("claim response contract mismatch")
        expires = value.get("expiresInSeconds")
        if isinstance(expires, bool) or not isinstance(expires, int) or not (1 <= expires <= 300):
            raise ClaimProtocolError("invalid signed URL lifetime")
        return cls(
            job_id=job_id,
            expires_in_seconds=expires,
            contract=contract,
            source_download_url=_https_url(value.get("sourceDownloadUrl")),
            target_download_url=_https_url(value.get("targetDownloadUrl")),
        )


def claim_local_repair(
    endpoint: str,
    ticket: LaunchTicket,
    identity: DeviceIdentity,
    transport: ClaimTransport,
) -> ClaimedRepair:
    endpoint = _https_url(endpoint)
    response = transport.post_json(endpoint, {
        "jobId": ticket.job_id,
        "claimToken": ticket.claim_token,
        "devicePublicKeySpki": identity.public_key_spki,
    })
    return ClaimedRepair.parse(response, expected_job_id=ticket.job_id)
