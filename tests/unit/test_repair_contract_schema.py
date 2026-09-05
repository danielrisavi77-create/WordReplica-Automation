import copy
import json
from pathlib import Path

import pytest

from word_replica.repair_contract.contract import RepairContractSchemaError, parse_repair_contract_v1

FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"
CONTRACT_PATH = FIXTURE_DIR / "valid-contract.json"


def _load_raw() -> dict:
    return json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))


def _mutate(raw: dict, path: list, value) -> dict:
    mutated = copy.deepcopy(raw)
    node = mutated
    for key in path[:-1]:
        node = node[key]
    node[path[-1]] = value
    return mutated


def test_parses_the_fixture_contract_without_touching_the_signature():
    contract = parse_repair_contract_v1(_load_raw())
    assert contract.contract_version == 1
    assert contract.engine_min_version == "0.1.0"
    assert contract.target_sha256 == "561a4ba09db7f3bc9a109ac02496321f8d866ececc338fc11a6fe30436977cb7"
    assert contract.target_size == 35
    assert contract.target_file_name == "Kalogjera - seminar Havel-popravljeno.docx"
    assert len(contract.requests) == 2
    assert len(contract.allowed_exceptions) == 1
    assert contract.verification_policy.required_gates == (
        "G0", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9",
    )


def test_parsed_signed_request_params_are_deeply_immutable_and_detached_from_raw():
    raw = _load_raw()
    contract = parse_repair_contract_v1(raw)

    raw["requests"][0]["params"]["fontName"] = "Tampered"
    raw["requests"][1]["params"]["levels"].append(99)

    assert contract.requests[0].params["fontName"] == "Times New Roman"
    assert contract.requests[1].params["levels"] == (1, 2)
    with pytest.raises(TypeError):
        contract.requests[0].params["fontName"] = "Tampered"


def test_rejects_unknown_top_level_key():
    raw = _mutate(_load_raw(), ["extraField"], "sneaky")
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


def test_rejects_missing_top_level_key():
    raw = _load_raw()
    del raw["userId"]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


@pytest.mark.parametrize("field", ["targetSha256", "targetSize", "targetFileName"])
def test_rejects_missing_signed_target_identity_field(field):
    raw = _load_raw()
    del raw[field]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


def test_rejects_non_lowercase_target_sha256():
    raw = _mutate(_load_raw(), ["targetSha256"], "A" * 64)
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-hash"


@pytest.mark.parametrize("value", [0, -1, True, 20 * 1024 * 1024 + 1])
def test_rejects_invalid_target_size(value):
    raw = _mutate(_load_raw(), ["targetSize"], value)
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


@pytest.mark.parametrize("value", ["../target.docx", "target.pdf", "CON.docx"])
def test_rejects_unsafe_target_filename(value):
    raw = _mutate(_load_raw(), ["targetFileName"], value)
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


def test_rejects_bool_as_int_source_size():
    raw = _mutate(_load_raw(), ["sourceSize"], True)
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


def test_rejects_unsafe_docx_filename_with_path_traversal():
    raw = _mutate(_load_raw(), ["sourceFileName"], "../evil.docx")
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-shape"


def test_rejects_filename_without_docx_extension():
    raw = _mutate(_load_raw(), ["sourceFileName"], "notadocx.pdf")
    with pytest.raises(RepairContractSchemaError):
        parse_repair_contract_v1(raw)


def test_rejects_reserved_windows_device_name():
    raw = _mutate(_load_raw(), ["sourceFileName"], "CON.docx")
    with pytest.raises(RepairContractSchemaError):
        parse_repair_contract_v1(raw)


def test_rejects_bad_uuid():
    raw = _mutate(_load_raw(), ["jobId"], "not-a-uuid")
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-id"


def test_rejects_bad_semver():
    raw = _mutate(_load_raw(), ["engineMinVersion"], "1.0")
    with pytest.raises(RepairContractSchemaError):
        parse_repair_contract_v1(raw)


def test_rejects_engine_max_below_engine_min():
    raw = copy.deepcopy(_load_raw())
    raw["engineMinVersion"] = "2.0.0"
    raw["engineMaxVersion"] = "1.0.0"
    with pytest.raises(RepairContractSchemaError):
        parse_repair_contract_v1(raw)


def test_rejects_expires_at_not_after_created_at():
    raw = copy.deepcopy(_load_raw())
    raw["expiresAt"] = raw["createdAt"]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "invalid-time"


def test_rejects_unknown_fixer_id():
    raw = _mutate(_load_raw(), ["requests", 0, "fixerId"], "not-a-real-fixer")
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "unknown-fixer"


def test_rejects_footer_page_fixer_as_a_standalone_request():
    raw = _mutate(_load_raw(), ["requests", 0, "fixerId"], "footer-page-fixer")
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "standalone-fixer-denied"


def test_rejects_duplicate_request_ids():
    raw = copy.deepcopy(_load_raw())
    raw["requests"][1]["requestId"] = raw["requests"][0]["requestId"]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "duplicate-request-id"


def test_rejects_too_many_requests():
    raw = copy.deepcopy(_load_raw())
    template = raw["requests"][0]
    raw["requests"] = [
        {**template, "requestId": f"req-{index:04d}"} for index in range(1, 66)
    ]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "request-count"


def test_rejects_orphan_exception_with_no_matching_request():
    raw = copy.deepcopy(_load_raw())
    raw["allowedExceptions"][0]["requestId"] = "req-9999"
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "orphan-exception"


def test_rejects_incomplete_required_gates():
    raw = copy.deepcopy(_load_raw())
    raw["verificationPolicy"]["requiredGates"] = ["G0", "G1"]
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "insufficient-verification-policy"


def test_rejects_verification_policy_flag_set_to_false():
    raw = copy.deepcopy(_load_raw())
    raw["verificationPolicy"]["requireOpenAndRepairFalse"] = False
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "insufficient-verification-policy"


def test_rejects_output_policy_overwrite_source_true():
    raw = copy.deepcopy(_load_raw())
    raw["outputPolicy"]["overwriteSource"] = True
    with pytest.raises(RepairContractSchemaError) as excinfo:
        parse_repair_contract_v1(raw)
    assert excinfo.value.code == "unsafe-output-policy"
