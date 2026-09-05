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
    working_output_path=r"C:\projects\project-1\output\working.docx",
    destination_path=r"C:\out\repaired.docx",
    engine_version="0.1.0",
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
        "projectId", "workingOutputPath", "destinationPath", "engineVersion",
    }
    assert data["schemaVersion"] == 2
    assert data["workingOutputPath"] != data["destinationPath"]


def test_validate_succeeds_when_every_bound_value_matches(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    store.create(binding)
    validated = store.validate(
        binding.job_id,
        contract_sha256=VALID_KWARGS["contract_sha256"],
        source_sha256=VALID_KWARGS["source_sha256"],
        target_sha256=VALID_KWARGS["target_sha256"],
        project_id=VALID_KWARGS["project_id"],
        working_output_path=VALID_KWARGS["working_output_path"],
        destination_path=VALID_KWARGS["destination_path"],
        engine_version=VALID_KWARGS["engine_version"],
    )
    assert validated == binding


@pytest.mark.parametrize("field,new_value", [
    ("contract_sha256", "z" * 64),
    ("source_sha256", "z" * 64),
    ("target_sha256", "z" * 64),
    ("project_id", "project-2"),
    ("working_output_path", r"C:\projects\project-1\output\different.docx"),
    ("destination_path", r"C:\out\different.docx"),
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
        "project_id": VALID_KWARGS["project_id"],
        "working_output_path": VALID_KWARGS["working_output_path"],
        "destination_path": VALID_KWARGS["destination_path"],
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
            project_id=VALID_KWARGS["project_id"],
            working_output_path=VALID_KWARGS["working_output_path"],
            destination_path=VALID_KWARGS["destination_path"],
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
    data["schemaVersion"] = 1
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


def test_build_rejects_aliased_working_and_destination_paths():
    with pytest.raises(ValueError):
        RepairRunBinding.build(
            **{
                **VALID_KWARGS,
                "destination_path": VALID_KWARGS["working_output_path"],
            }
        )


def test_store_is_idempotent_for_same_binding_but_refuses_rebinding(tmp_path):
    store = RepairRunBindingStore(app_root=tmp_path)
    binding = RepairRunBinding.build(**VALID_KWARGS)
    first = store.create(binding)
    assert store.create(binding) == first
    changed = RepairRunBinding.build(**{**VALID_KWARGS, "project_id": "project-2"})
    with pytest.raises(RepairPackageError) as excinfo:
        store.create(changed)
    assert excinfo.value.code == "binding-mismatch"
