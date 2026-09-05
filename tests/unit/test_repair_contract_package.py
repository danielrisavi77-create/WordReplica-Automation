import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

import pytest

from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.package import RepairPackageRequest, load_and_validate_package

FIXTURE_DIR = Path(__file__).resolve().parents[1] / "fixtures" / "repair_contract_v1"
CONTRACT_PATH = FIXTURE_DIR / "valid-contract.json"
PUBLIC_KEY_PATH = FIXTURE_DIR / "public-key.spki.b64url"

# The published fixture signs over this exact 28-byte public source vector.
SOURCE_BYTES = "PK-public-repair-contract-v1".encode("utf-8")
TARGET_BYTES = "PK-public-repair-contract-v1-target".encode("utf-8")
NOW = datetime(2026, 8, 16, 10, 30, tzinfo=timezone.utc)  # inside the fixture's 10:00-11:00 window
ENGINE_VERSION = "0.1.0"  # matches the fixture's engineMinVersion == engineMaxVersion


def _write_package(tmp_path: Path, *, source_bytes: bytes = SOURCE_BYTES, target_bytes: bytes = TARGET_BYTES) -> RepairPackageRequest:
    contract = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
    original_dir = tmp_path / "in"
    original_dir.mkdir()
    original_path = original_dir / contract["sourceFileName"]
    original_path.write_bytes(source_bytes)

    target_path = original_dir / contract["targetFileName"]
    target_path.write_bytes(target_bytes)

    contract_path = tmp_path / "contract.json"
    shutil.copy(CONTRACT_PATH, contract_path)
    public_key_path = tmp_path / "public-key.spki.b64url"
    shutil.copy(PUBLIC_KEY_PATH, public_key_path)

    return RepairPackageRequest(
        original_path=original_path,
        target_path=target_path,
        contract_path=contract_path,
        public_key_path=public_key_path,
        output_dir=tmp_path / "out",
    )


def test_validates_a_correct_package_and_reserves_an_output_path(tmp_path):
    request = _write_package(tmp_path)
    validated = load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert validated.original_snapshot.sha256 == validated.contract.source_sha256
    assert validated.target_snapshot.sha256 == validated.contract.target_sha256
    assert validated.target_snapshot.size == validated.contract.target_size
    assert validated.output_path.name == validated.contract.output_policy.suggested_file_name
    assert validated.output_path.parent == request.output_dir.resolve()
    assert not validated.output_path.exists()


def test_never_touches_word_on_any_preflight_error(tmp_path):
    word_started = {"value": False}

    def start_word():
        word_started["value"] = True

    request = _write_package(tmp_path, source_bytes=b"tampered bytes")
    with pytest.raises(RepairPackageError):
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert word_started["value"] is False  # start_word was never called


def test_rejects_aliased_source_and_target_paths(tmp_path):
    request = _write_package(tmp_path)
    aliased = RepairPackageRequest(
        original_path=request.original_path,
        target_path=request.original_path,
        contract_path=request.contract_path,
        public_key_path=request.public_key_path,
        output_dir=request.output_dir,
    )
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(aliased, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "aliased-paths"


def test_rejects_wrong_original_filename(tmp_path):
    request = _write_package(tmp_path)
    renamed = request.original_path.with_name("wrong-name.docx")
    request.original_path.rename(renamed)
    request = RepairPackageRequest(renamed, request.target_path, request.contract_path, request.public_key_path, request.output_dir)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "filename-mismatch"


def test_rejects_wrong_target_filename(tmp_path):
    request = _write_package(tmp_path)
    renamed = request.target_path.with_name("wrong-name.docx")
    request.target_path.rename(renamed)
    request = RepairPackageRequest(request.original_path, renamed, request.contract_path, request.public_key_path, request.output_dir)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "filename-mismatch"


def test_rejects_source_bytes_that_do_not_match_the_signed_hash(tmp_path):
    request = _write_package(tmp_path, source_bytes=b"not the signed source bytes!")
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "source-hash-mismatch"


def test_rejects_target_bytes_that_do_not_match_the_signed_hash_before_word(tmp_path):
    tampered_same_size = bytes([TARGET_BYTES[0] ^ 1]) + TARGET_BYTES[1:]
    request = _write_package(tmp_path, target_bytes=tampered_same_size)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "target-hash-mismatch"


def test_rejects_target_size_mismatch_before_hash_check(tmp_path):
    request = _write_package(tmp_path, target_bytes=TARGET_BYTES + b"!")
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "target-size-mismatch"


def test_rejects_before_created_at(tmp_path):
    request = _write_package(tmp_path)
    too_early = datetime(2026, 8, 16, 9, 0, tzinfo=timezone.utc)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=too_early, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "invalid-time"


def test_rejects_expired_contract(tmp_path):
    request = _write_package(tmp_path)
    too_late = datetime(2026, 8, 16, 12, 0, tzinfo=timezone.utc)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=too_late, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "expired"


def test_rejects_engine_version_below_range(tmp_path):
    request = _write_package(tmp_path)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version="0.9.9")
    assert excinfo.value.code == "engine-out-of-range"


def test_rejects_engine_version_above_range(tmp_path):
    request = _write_package(tmp_path)
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version="1.0.1")
    assert excinfo.value.code == "engine-out-of-range"


def test_rejects_missing_original_file(tmp_path):
    request = _write_package(tmp_path)
    request.original_path.unlink()
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "missing-file"


def test_rejects_output_dir_that_is_actually_a_file(tmp_path):
    request = _write_package(tmp_path)
    request.output_dir.parent.mkdir(parents=True, exist_ok=True)
    request.output_dir.write_bytes(b"not a directory")
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "invalid-output-dir"


def test_reserves_a_collision_safe_suffixed_name_when_output_already_exists(tmp_path):
    request = _write_package(tmp_path)
    request.output_dir.mkdir(parents=True, exist_ok=True)
    validated_contract = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
    suggested = validated_contract["outputPolicy"]["suggestedFileName"]
    (request.output_dir / suggested).write_bytes(b"already there")

    validated = load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert validated.output_path.name != suggested
    assert validated.output_path.stem.startswith(Path(suggested).stem)
    assert not validated.output_path.exists()


def test_rejects_tampered_contract_signature_before_any_file_check(tmp_path):
    request = _write_package(tmp_path)
    contract = json.loads(request.contract_path.read_text(encoding="utf-8"))
    contract["sourceSha256"] = "0" * 64
    request.contract_path.write_text(json.dumps(contract), encoding="utf-8")
    with pytest.raises(RepairPackageError) as excinfo:
        load_and_validate_package(request, now=NOW, engine_version=ENGINE_VERSION)
    assert excinfo.value.code == "signature-mismatch"
