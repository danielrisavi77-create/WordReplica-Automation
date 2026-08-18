import json

import pytest

from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.binding import RepairRunBinding, RepairRunBindingStore

VALID_KWARGS = dict(
    job_id="11111111-1111-4111-8111-111111111111",
    contract_sha256="a" * 64,
    source_sha256="b" * 64,
    target_sha256="c" * 64,
    project_id="project-1",
    output_path=r"C:\out\repaired.docx",
    engine_version="1.0.0",
)


def test_build_round_trips_through_json():
    binding = RepairRunBinding.build(**VALID_KWARGS)
    restored = RepairRunBinding.from_json(binding.to_json())
    assert restored == binding


def test_store_creates_and_loads_a_binding_atomically(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    path = store.create(binding)
    assert path.is_file()
    assert not path.with_suffix(".json.tmp").exists()
    loaded = store.load(binding.job_id)
    assert loaded == binding


def test_stored_binding_json_has_exact_schema_keys_and_no_document_text(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    path = store.create(binding)
    data = json.loads(path.read_text(encoding="utf-8"))
    assert set(data.keys()) == {
        "schemaVersion", "jobId", "contractSha256", "sourceSha256", "targetSha256",
        "projectId", "outputPath", "engineVersion",
    }
    assert data["schemaVersion"] == 1


def test_validate_succeeds_when_every_bound_value_matches(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    store.create(binding)
    validated = store.validate(
        binding.job_id,
        contract_sha256=VALID_KWARGS["contract_sha256"],
        source_sha256=VALID_KWARGS["source_sha256"],
        target_sha256=VALID_KWARGS["target_sha256"],
        output_path=VALID_KWARGS["output_path"],
        engine_version=VALID_KWARGS["engine_version"],
    )
    assert validated == binding


@pytest.mark.parametrize("field,new_value", [
    ("contract_sha256", "z" * 64),
    ("source_sha256", "z" * 64),
    ("target_sha256", "z" * 64),
    ("output_path", r"C:\out\different.docx"),
    ("engine_version", "2.0.0"),
])
def test_validate_rejects_any_drifted_bound_value(tmp_path, field, new_value):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    store.create(binding)
    call_kwargs = {
        "contract_sha256": VALID_KWARGS["contract_sha256"],
        "source_sha256": VALID_KWARGS["source_sha256"],
        "target_sha256": VALID_KWARGS["target_sha256"],
        "output_path": VALID_KWARGS["output_path"],
        "engine_version": VALID_KWARGS["engine_version"],
    }
    call_kwargs[field] = new_value
    with pytest.raises(RepairPackageError) as excinfo:
        store.validate(binding.job_id, **call_kwargs)
    assert excinfo.value.code == "binding-mismatch"
    assert field in str(excinfo.value)


def test_validate_rejects_a_binding_filed_under_the_wrong_job_id(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    path = store.create(binding)
    # Simulate a binding file copied into another job's folder.
    other_job_dir = store.jobs_dir / "22222222-2222-4222-8222-222222222222"
    other_job_dir.mkdir(parents=True)
    (other_job_dir / "binding.json").write_text(path.read_text(encoding="utf-8"), encoding="utf-8")

    with pytest.raises(RepairPackageError) as excinfo:
        store.validate(
            "22222222-2222-4222-8222-222222222222",
            contract_sha256=VALID_KWARGS["contract_sha256"],
            source_sha256=VALID_KWARGS["source_sha256"],
            target_sha256=VALID_KWARGS["target_sha256"],
            output_path=VALID_KWARGS["output_path"],
            engine_version=VALID_KWARGS["engine_version"],
        )
    assert excinfo.value.code == "binding-mismatch"
    assert "job_id" in str(excinfo.value)


def test_load_missing_binding_fails_closed(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    with pytest.raises(RepairPackageError) as excinfo:
        store.load("does-not-exist")
    assert excinfo.value.code == "missing-binding"


def test_load_rejects_a_binding_with_wrong_schema_version(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    path = store.create(binding)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["schemaVersion"] = 2
    path.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(RepairPackageError) as excinfo:
        store.load(binding.job_id)
    assert excinfo.value.code == "invalid-binding"


def test_load_rejects_a_binding_with_unexpected_keys(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    path = store.create(binding)
    data = json.loads(path.read_text(encoding="utf-8"))
    data["unexpectedKey"] = "sneaky"
    path.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(RepairPackageError) as excinfo:
        store.load(binding.job_id)
    assert excinfo.value.code == "invalid-binding"
