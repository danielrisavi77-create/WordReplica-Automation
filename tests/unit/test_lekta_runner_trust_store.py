from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path

import pytest
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.serialization import (
    Encoding,
    PublicFormat,
    load_der_public_key,
    load_pem_public_key,
)

from word_replica.repair_contract.signature import decode_spki
from word_replica.runner import trust_store as trust_store_module
from word_replica.runner.trust_store import RunnerTrustError, load_trust_keys


FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"


def test_trust_store_loads_only_explicit_key_id_to_der_mapping(tmp_path: Path) -> None:
    spki = (FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8").strip()
    path = tmp_path / "trusted_keys.json"
    path.write_text(json.dumps({
        "version": 1,
        "keys": [{"keyId": "lekta-prod-2026-01", "spkiBase64Url": spki}],
    }), encoding="utf-8")

    assert load_trust_keys(path) == {"lekta-prod-2026-01": decode_spki(spki)}


@pytest.mark.parametrize("payload", [
    {"version": 1, "keys": []},
    {"version": 2, "keys": []},
    {"version": 1, "keys": [{"keyId": "x", "spkiBase64Url": "abc", "privateKey": "no"}]},
    {"version": 1, "keys": [
        {"keyId": "same", "spkiBase64Url": "abc"},
        {"keyId": "same", "spkiBase64Url": "abc"},
    ]},
])
def test_trust_store_rejects_empty_malformed_or_duplicate_keys(tmp_path: Path, payload: dict) -> None:
    path = tmp_path / "trusted_keys.json"
    path.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(RunnerTrustError):
        load_trust_keys(path)


def test_release_build_prepares_a_public_only_trust_store(tmp_path: Path) -> None:
    prepare = getattr(trust_store_module, "prepare_release_trust_store", None)
    assert callable(prepare), "release trust-store preparation is missing"

    spki = (FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8").strip()
    destination = tmp_path / "generated" / "trusted_keys.json"

    result = prepare(
        public_key_path=FIXTURE_DIR / "public-key.spki.b64url",
        key_id="lekta-prod-test",
        destination=destination,
    )

    assert result.path == destination
    assert result.contract_public_key_sha256 == sha256(decode_spki(spki)).hexdigest()
    assert json.loads(destination.read_text(encoding="utf-8")) == {
        "version": 1,
        "keys": [{"keyId": "lekta-prod-test", "spkiBase64Url": spki}],
    }
    assert load_trust_keys(destination) == {
        "lekta-prod-test": decode_spki(spki),
    }


def test_release_key_fingerprint_is_canonical_der_spki_sha256(tmp_path: Path) -> None:
    spki_text = (FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8").strip()
    der = decode_spki(spki_text)
    prepared = trust_store_module.prepare_release_trust_store(
        public_key_path=FIXTURE_DIR / "public-key.spki.b64url",
        key_id="lekta-prod-test",
        destination=tmp_path / "trusted_keys.json",
    )

    assert prepared.contract_public_key_sha256 == sha256(der).hexdigest()
    assert prepared.path == tmp_path / "trusted_keys.json"


def test_fingerprint_ignores_pem_text_format_but_changes_for_another_key() -> None:
    fixture_der = decode_spki(
        (FIXTURE_DIR / "public-key.spki.b64url").read_text(encoding="utf-8").strip()
    )
    fixture_key = load_der_public_key(fixture_der)
    pem_lf = fixture_key.public_bytes(Encoding.PEM, PublicFormat.SubjectPublicKeyInfo)
    pem_crlf = pem_lf.replace(b"\n", b"\r\n")
    fingerprints = []
    for pem in (pem_lf, pem_crlf):
        parsed = load_pem_public_key(pem)
        canonical_der = parsed.public_bytes(Encoding.DER, PublicFormat.SubjectPublicKeyInfo)
        fingerprints.append(trust_store_module.canonical_p256_spki_sha256(canonical_der))

    other_der = ec.generate_private_key(ec.SECP256R1()).public_key().public_bytes(
        Encoding.DER,
        PublicFormat.SubjectPublicKeyInfo,
    )
    assert fingerprints == [sha256(fixture_der).hexdigest()] * 2
    assert trust_store_module.canonical_p256_spki_sha256(other_der) != fingerprints[0]
    assert sha256(pem_lf).hexdigest() != fingerprints[0]
