"""Pinned public-key trust store for signed Lekta Repair Contracts."""
from __future__ import annotations

import json
import os
import re
import tempfile
from pathlib import Path

from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.serialization import load_der_public_key

from word_replica.repair_contract.signature import decode_spki


_KEY_ID = re.compile(r"^[A-Za-z0-9._-]{1,80}$")
_MAX_TRUST_STORE_BYTES = 64 * 1024


class RunnerTrustError(ValueError):
    """The executable's pinned contract trust store is absent or invalid."""


def load_trust_keys(path: Path) -> dict[str, bytes]:
    path = Path(path)
    try:
        if path.stat().st_size > _MAX_TRUST_STORE_BYTES:
            raise RunnerTrustError("runner trust store is too large")
        raw = json.loads(path.read_text(encoding="utf-8"))
    except RunnerTrustError:
        raise
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RunnerTrustError("runner trust store is unavailable") from exc

    if not isinstance(raw, dict) or set(raw) != {"version", "keys"} or raw.get("version") != 1:
        raise RunnerTrustError("invalid runner trust store")
    keys = raw.get("keys")
    if not isinstance(keys, list) or not keys:
        raise RunnerTrustError("runner trust store has no keys")

    trusted: dict[str, bytes] = {}
    for item in keys:
        if not isinstance(item, dict) or set(item) != {"keyId", "spkiBase64Url"}:
            raise RunnerTrustError("invalid runner trust key")
        key_id = item.get("keyId")
        encoded = item.get("spkiBase64Url")
        if not isinstance(key_id, str) or not _KEY_ID.fullmatch(key_id) or key_id in trusted:
            raise RunnerTrustError("invalid or duplicate runner trust key id")
        if not isinstance(encoded, str):
            raise RunnerTrustError("invalid runner trust public key")
        try:
            der = decode_spki(encoded)
            public_key = load_der_public_key(der)
        except Exception as exc:
            raise RunnerTrustError("invalid runner trust public key") from exc
        if not isinstance(public_key, ec.EllipticCurvePublicKey) or not isinstance(public_key.curve, ec.SECP256R1):
            raise RunnerTrustError("runner trust key is not P-256")
        trusted[key_id] = der
    return trusted


def prepare_release_trust_store(
    *,
    public_key_path: Path,
    key_id: str,
    destination: Path,
) -> Path:
    """Atomically prepare the public-only trust asset for a portable release."""
    if not isinstance(key_id, str) or not _KEY_ID.fullmatch(key_id):
        raise RunnerTrustError("invalid runner trust key id")
    public_key_path = Path(public_key_path)
    destination = Path(destination)
    try:
        spki = public_key_path.read_text(encoding="utf-8").strip()
    except (OSError, UnicodeDecodeError) as exc:
        raise RunnerTrustError("runner trust public key is unavailable") from exc

    payload = json.dumps(
        {
            "version": 1,
            "keys": [{"keyId": key_id, "spkiBase64Url": spki}],
        },
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix="trusted-keys-", suffix=".json.tmp", dir=destination.parent,
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as target:
            target.write(payload)
            target.flush()
            os.fsync(target.fileno())
        load_trust_keys(temporary)
        os.replace(temporary, destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    return destination
