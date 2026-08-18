import copy
import json
from pathlib import Path

import pytest

from word_replica.repair_contract.signature import (
    RepairContractSignatureError,
    decode_spki,
    verify_signed_contract,
)

FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"
CONTRACT_PATH = FIXTURE_DIR / "valid-contract.json"
PUBLIC_KEY_PATH = FIXTURE_DIR / "public-key.spki.b64url"


def _load_raw() -> dict:
    return json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))


def _trusted_keys() -> dict:
    return {"fixture-2026-08-16": decode_spki(PUBLIC_KEY_PATH.read_text().strip())}


def _mutate(raw: dict, path: list, value) -> dict:
    mutated = copy.deepcopy(raw)
    node = mutated
    for key in path[:-1]:
        node = node[key]
    node[path[-1]] = value
    return mutated


def test_verifies_the_published_lekta_fixture_contract():
    contract = verify_signed_contract(_load_raw(), _trusted_keys())
    assert contract.source_file_name == "Kalogjera - seminar Havel.docx"
    assert contract.output_policy.suggested_file_name == "Kalogjera - seminar Havel-popravljeno.docx"
    assert contract.job_id == "11111111-1111-4111-8111-111111111111"
    assert [request.fixer_id for request in contract.requests] == ["font-fixer", "heading-case-fixer"]


def test_rejects_unknown_key_id():
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(_load_raw(), {})
    assert excinfo.value.code == "invalid-key-id"


def test_rejects_unsupported_algorithm():
    raw = _mutate(_load_raw(), ["contractSignature", "algorithm"], "RS256")
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(raw, _trusted_keys())
    assert excinfo.value.code == "unsupported-algorithm"


def test_rejects_malformed_signature_encoding():
    raw = _mutate(_load_raw(), ["contractSignature", "value"], "not-base64url!!")
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(raw, _trusted_keys())
    assert excinfo.value.code == "invalid-signature-encoding"


def test_rejects_a_63_byte_signature():
    # 63 bytes -> 84 base64url characters without padding.
    short = "A" * 84
    raw = _mutate(_load_raw(), ["contractSignature", "value"], short)
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(raw, _trusted_keys())
    assert excinfo.value.code == "invalid-signature-encoding"


@pytest.mark.parametrize("path,value", [
    (["sourceSha256"], "0" * 64),
    (["sourceSize"], 999999),
    (["sourceFileName"], "different-file-name.docx"),
    (["requests", 0, "params", "fontSizePt"], 99),
    (["expiresAt"], "2099-01-01T00:00:00.000Z"),
    (["outputPolicy", "suggestedFileName"], "other-name.docx"),
    (["jobId"], "99999999-9999-4999-8999-999999999999"),
])
def test_any_signed_field_mutation_breaks_the_signature(path, value):
    mutated = _mutate(_load_raw(), path, value)
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(mutated, _trusted_keys())
    assert excinfo.value.code == "signature-mismatch"


def test_reordering_the_requests_array_breaks_the_signature():
    raw = _load_raw()
    mutated = copy.deepcopy(raw)
    mutated["requests"] = list(reversed(mutated["requests"]))
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(mutated, _trusted_keys())
    assert excinfo.value.code == "signature-mismatch"


def test_an_extra_unsigned_top_level_key_breaks_the_signature():
    raw = _mutate(_load_raw(), ["extraField"], "sneaky")
    with pytest.raises(RepairContractSignatureError) as excinfo:
        verify_signed_contract(raw, _trusted_keys())
    assert excinfo.value.code == "signature-mismatch"
