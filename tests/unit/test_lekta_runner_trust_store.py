from __future__ import annotations

import json
from pathlib import Path

import pytest

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

    assert result == destination
    assert json.loads(destination.read_text(encoding="utf-8")) == {
        "version": 1,
        "keys": [{"keyId": "lekta-prod-test", "spkiBase64Url": spki}],
    }
    assert load_trust_keys(destination) == {
        "lekta-prod-test": decode_spki(spki),
    }
