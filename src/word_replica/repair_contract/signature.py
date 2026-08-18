"""ES256-P1363 signature verification for a Lekta Repair Contract v1.

Mirrors src/repair/contract/signature.ts and canonical-json.ts on the Lekta
side. The wire signature is raw IEEE P1363 (r || s, 64 bytes) base64url with
no padding; Python's `cryptography` verifies DER-encoded ECDSA signatures, so
the P1363 bytes are re-encoded to DER only for the local verify call — the
wire format itself is never touched.
"""
from __future__ import annotations

import base64
import re
from collections.abc import Mapping
from typing import Any

import rfc8785
from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import encode_dss_signature
from cryptography.hazmat.primitives.hashes import SHA256
from cryptography.hazmat.primitives.serialization import load_der_public_key

from word_replica.repair_contract.contract import RepairContractV1, parse_repair_contract_v1

SUPPORTED_ALGORITHM = "ES256-P1363"
_KEY_ID_RE = re.compile(r"^[A-Za-z0-9._-]{1,80}$")


class RepairContractSignatureError(ValueError):
    def __init__(self, code: str, message: str = "") -> None:
        super().__init__(f"{code}: {message}" if message else code)
        self.code = code


def decode_spki(b64url_text: str) -> bytes:
    """Decode a base64url-no-padding SPKI public key into DER bytes."""
    text = b64url_text.strip()
    padded = text + "=" * (-len(text) % 4)
    try:
        return base64.urlsafe_b64decode(padded)
    except Exception as exc:
        raise RepairContractSignatureError("invalid-signature-encoding", "malformed public key") from exc


def _decode_base64url_no_padding(text: str) -> bytes:
    if not isinstance(text, str) or "=" in text or "+" in text or "/" in text:
        raise RepairContractSignatureError("invalid-signature-encoding", "not canonical base64url")
    padded = text + "=" * (-len(text) % 4)
    try:
        return base64.urlsafe_b64decode(padded)
    except Exception as exc:
        raise RepairContractSignatureError("invalid-signature-encoding", "malformed base64url") from exc


def canonical_unsigned_bytes(raw: Mapping[str, Any]) -> bytes:
    """UTF-8 bytes of the RFC 8785 (JCS) canonical form, minus contractSignature.

    JCS matches the ECMAScript-based canonicalization Lekta's TS signer uses
    (sorted object keys, ECMA-262 number-to-string, dense-array-only arrays),
    so this reproduces byte-identical signed payloads without depending on
    Lekta's own implementation.
    """
    unsigned = {key: value for key, value in raw.items() if key != "contractSignature"}
    return rfc8785.dumps(unsigned)


def verify_signed_contract(raw: Mapping[str, Any], trusted_keys: Mapping[str, bytes]) -> RepairContractV1:
    """Verify a Lekta Repair Contract v1's signature, then strictly parse it.

    `trusted_keys` maps a `keyId` to its DER-encoded SPKI public key bytes.
    Fails closed: any error raises RepairContractSignatureError (or the
    schema parser's RepairContractSchemaError) before a contract is returned.
    """
    if not isinstance(raw, dict):
        raise RepairContractSignatureError("invalid-shape", "contract must be a JSON object")
    envelope = raw.get("contractSignature")
    if not isinstance(envelope, dict):
        raise RepairContractSignatureError("invalid-shape", "missing contractSignature")

    algorithm = envelope.get("algorithm")
    if algorithm != SUPPORTED_ALGORITHM:
        raise RepairContractSignatureError("unsupported-algorithm", str(algorithm))

    key_id = envelope.get("keyId")
    if not isinstance(key_id, str) or not _KEY_ID_RE.match(key_id):
        raise RepairContractSignatureError("invalid-key-id", str(key_id))

    public_key_der = trusted_keys.get(key_id)
    if public_key_der is None:
        raise RepairContractSignatureError("invalid-key-id", f"unknown key id: {key_id}")

    signature_value = envelope.get("value")
    if not isinstance(signature_value, str):
        raise RepairContractSignatureError("invalid-signature-encoding", "value must be a string")
    signature_bytes = _decode_base64url_no_padding(signature_value)
    if len(signature_bytes) != 64:
        raise RepairContractSignatureError("invalid-signature-encoding", f"expected 64 bytes, got {len(signature_bytes)}")

    r = int.from_bytes(signature_bytes[:32], "big")
    s = int.from_bytes(signature_bytes[32:], "big")
    der_signature = encode_dss_signature(r, s)

    try:
        public_key = load_der_public_key(public_key_der)
    except Exception as exc:
        raise RepairContractSignatureError("invalid-key-id", "malformed trusted public key") from exc
    if not isinstance(public_key, ec.EllipticCurvePublicKey) or not isinstance(public_key.curve, ec.SECP256R1):
        raise RepairContractSignatureError("invalid-key-id", "trusted key is not P-256")

    payload = canonical_unsigned_bytes(raw)
    try:
        public_key.verify(der_signature, payload, ec.ECDSA(SHA256()))
    except InvalidSignature as exc:
        raise RepairContractSignatureError("signature-mismatch") from exc

    return parse_repair_contract_v1(raw)
