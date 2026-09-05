"""Durable job-to-resume binding for a repair-package run.

Stored under the WordReplica app root at repair_jobs/<job_id>/binding.json.
Written atomically (temp file then replace) so a crash mid-write never
leaves a half-written binding. Resume must revalidate every bound value
before Word reopens; any drift (a changed contract, original, target,
project, working/output paths, engine version, or a binding filed under the wrong job)
fails closed rather than resuming against inputs that no longer match
what was originally verified.

Never stores document text, an entitlement token, or a private key —
only hashes, the job/project identifiers and the two distinct output paths.
"""
from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path

from platformdirs import user_documents_dir

from word_replica.domain.errors import RepairPackageError

SCHEMA_VERSION = 2
_FIELDS = (
    "schemaVersion", "jobId", "contractSha256", "sourceSha256", "targetSha256",
    "projectId", "workingOutputPath", "destinationPath", "engineVersion",
)


@dataclass(frozen=True, slots=True)
class RepairRunBinding:
    schema_version: int
    job_id: str
    contract_sha256: str
    source_sha256: str
    target_sha256: str
    project_id: str
    working_output_path: str
    destination_path: str
    engine_version: str

    @classmethod
    def build(
        cls,
        *,
        job_id: str,
        contract_sha256: str,
        source_sha256: str,
        target_sha256: str,
        project_id: str,
        working_output_path: str | Path,
        destination_path: str | Path,
        engine_version: str,
    ) -> "RepairRunBinding":
        working = str(Path(working_output_path))
        destination = str(Path(destination_path))
        if os.path.normcase(os.path.abspath(working)) == os.path.normcase(os.path.abspath(destination)):
            raise ValueError("working output and destination must be different paths")
        return cls(
            schema_version=SCHEMA_VERSION,
            job_id=job_id,
            contract_sha256=contract_sha256,
            source_sha256=source_sha256,
            target_sha256=target_sha256,
            project_id=project_id,
            working_output_path=working,
            destination_path=destination,
            engine_version=engine_version,
        )

    def to_json(self) -> dict:
        data = asdict(self)
        return {
            "schemaVersion": data["schema_version"],
            "jobId": data["job_id"],
            "contractSha256": data["contract_sha256"],
            "sourceSha256": data["source_sha256"],
            "targetSha256": data["target_sha256"],
            "projectId": data["project_id"],
            "workingOutputPath": data["working_output_path"],
            "destinationPath": data["destination_path"],
            "engineVersion": data["engine_version"],
        }

    @classmethod
    def from_json(cls, data: dict) -> "RepairRunBinding":
        if not isinstance(data, dict) or set(data.keys()) != set(_FIELDS):
            raise ValueError("binding.json has an unexpected shape")
        if data["schemaVersion"] != SCHEMA_VERSION:
            raise ValueError(f"unsupported binding schema version: {data['schemaVersion']!r}")
        return cls.build(
            job_id=data["jobId"],
            contract_sha256=data["contractSha256"],
            source_sha256=data["sourceSha256"],
            target_sha256=data["targetSha256"],
            project_id=data["projectId"],
            working_output_path=data["workingOutputPath"],
            destination_path=data["destinationPath"],
            engine_version=data["engineVersion"],
        )


class RepairRunBindingStore:
    def __init__(self, app_root: Path | None = None) -> None:
        self.app_root = app_root or Path(user_documents_dir()) / "WordReplica"
        self.jobs_dir = self.app_root / "repair_jobs"

    def _binding_path(self, job_id: str) -> Path:
        return self.jobs_dir / job_id / "binding.json"

    def create(self, binding: RepairRunBinding) -> Path:
        path = self._binding_path(binding.job_id)
        if path.exists():
            existing = self.load(binding.job_id)
            if existing == binding:
                return path
            raise RepairPackageError("binding-mismatch", "refusing to replace an existing job binding")
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = path.with_suffix(".json.tmp")
        tmp_path.write_text(json.dumps(binding.to_json(), indent=2, sort_keys=True), encoding="utf-8")
        tmp_path.replace(path)
        return path

    def load(self, job_id: str) -> RepairRunBinding:
        path = self._binding_path(job_id)
        if not path.is_file():
            raise RepairPackageError("missing-binding", str(path))
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except ValueError as exc:
            raise RepairPackageError("invalid-binding", str(exc)) from exc
        try:
            return RepairRunBinding.from_json(data)
        except ValueError as exc:
            raise RepairPackageError("invalid-binding", str(exc)) from exc

    def validate(
        self,
        job_id: str,
        *,
        contract_sha256: str,
        source_sha256: str,
        target_sha256: str,
        project_id: str,
        working_output_path: str | Path,
        destination_path: str | Path,
        engine_version: str,
    ) -> RepairRunBinding:
        """Load the stored binding for job_id and fail closed on any drift.

        Resume must call this — never RepairRunBindingStore.load directly —
        so a changed contract, original, target, output path or engine
        version is rejected before Word reopens.
        """
        binding = self.load(job_id)
        mismatched: list[str] = []
        if binding.job_id != job_id:
            mismatched.append("job_id")
        if binding.contract_sha256 != contract_sha256:
            mismatched.append("contract_sha256")
        if binding.source_sha256 != source_sha256:
            mismatched.append("source_sha256")
        if binding.target_sha256 != target_sha256:
            mismatched.append("target_sha256")
        if binding.project_id != project_id:
            mismatched.append("project_id")
        if binding.working_output_path != str(Path(working_output_path)):
            mismatched.append("working_output_path")
        if binding.destination_path != str(Path(destination_path)):
            mismatched.append("destination_path")
        if binding.engine_version != engine_version:
            mismatched.append("engine_version")
        if mismatched:
            raise RepairPackageError("binding-mismatch", ",".join(mismatched))
        return binding
