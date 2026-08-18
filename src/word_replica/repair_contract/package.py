"""Explicit-path preflight for a local Lekta repair package.

Real difference from the original plan: Lekta's actual signed Repair
Contract v1 binds only the *source* identity (sourceSha256/sourceSize/
sourceFileName) plus an operation list — it does not sign a target hash,
because the contract is created before the fixers run and the corrected
bytes do not exist yet. So the target document's identity is NOT
cryptographically verified here; it is trusted by explicit local path plus
a filename match against the contract's own outputPolicy.suggestedFileName.
Only the original source has a real cryptographic integrity guarantee.
This must stay honestly reflected in the completion report (Task 5), not
papered over as if it were an equivalent check.

Nothing in this module touches Microsoft Word or any COM object; every
failure here happens before Word could ever be started.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path
from typing import Any

from word_replica.domain.errors import RepairPackageError
from word_replica.repair_contract.contract import RepairContractSchemaError, RepairContractV1
from word_replica.repair_contract.signature import RepairContractSignatureError, decode_spki, verify_signed_contract
from word_replica.services.source_guard import SourceSnapshot, capture_source

DEFAULT_MAX_LIFETIME_SECONDS = 24 * 60 * 60


@dataclass(frozen=True, slots=True)
class RepairPackageRequest:
    original_path: Path
    target_path: Path
    contract_path: Path
    public_key_path: Path
    output_dir: Path


@dataclass(frozen=True, slots=True)
class ValidatedRepairPackage:
    contract: RepairContractV1
    contract_sha256: str
    original_snapshot: SourceSnapshot
    target_snapshot: SourceSnapshot
    output_path: Path


def _same_path(a: Path, b: Path) -> bool:
    return os.path.normcase(str(a)) == os.path.normcase(str(b))


def _read_json(path: Path, code: str) -> Any:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise RepairPackageError(code, f"{path}: {exc}") from exc
    try:
        return json.loads(text)
    except ValueError as exc:
        raise RepairPackageError(code, f"{path}: {exc}") from exc


def _parse_iso8601_utc(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%dT%H:%M:%S.%fZ").replace(tzinfo=timezone.utc)


def _semver_tuple(value: str) -> tuple[int, int, int]:
    major, minor, patch = value.split(".")
    return int(major), int(minor), int(patch)


def reserve_output_path(output_dir: Path, suggested_file_name: str, *, avoid: set[Path]) -> Path:
    """Pick a non-overwriting output path under output_dir, starting from the
    signed suggested filename. Never returns a path that aliases `avoid`."""
    stem = Path(suggested_file_name).stem
    suffix = Path(suggested_file_name).suffix
    avoid_normalized = {os.path.normcase(str(path)) for path in avoid}

    candidate = output_dir / suggested_file_name
    counter = 2
    while candidate.exists() or os.path.normcase(str(candidate)) in avoid_normalized:
        candidate = output_dir / f"{stem} ({counter}){suffix}"
        counter += 1
    return candidate


def load_and_validate_package(
    request: RepairPackageRequest,
    *,
    now: datetime,
    engine_version: str,
    max_lifetime_seconds: int = DEFAULT_MAX_LIFETIME_SECONDS,
) -> ValidatedRepairPackage:
    original_path = request.original_path.resolve()
    target_path = request.target_path.resolve()
    contract_path = request.contract_path.resolve()
    public_key_path = request.public_key_path.resolve()
    output_dir = request.output_dir.resolve()

    for label, path in (
        ("original", original_path), ("target", target_path),
        ("contract", contract_path), ("public-key", public_key_path),
    ):
        if not path.is_file():
            raise RepairPackageError("missing-file", f"{label}: {path}")

    if _same_path(original_path, target_path):
        raise RepairPackageError("aliased-paths", "original and target must resolve to different files")
    if original_path.suffix.lower() != ".docx" or target_path.suffix.lower() != ".docx":
        raise RepairPackageError("invalid-shape", "original and target must be .docx files")

    raw_contract = _read_json(contract_path, "invalid-contract-json")
    if not isinstance(raw_contract, dict):
        raise RepairPackageError("invalid-contract-json", "contract must be a JSON object")
    envelope = raw_contract.get("contractSignature")
    key_id = envelope.get("keyId") if isinstance(envelope, dict) else None
    if not isinstance(key_id, str):
        raise RepairPackageError("invalid-contract-json", "missing contractSignature.keyId")

    try:
        public_key_der = decode_spki(public_key_path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise RepairPackageError("invalid-public-key", str(exc)) from exc
    except RepairContractSignatureError as exc:
        raise RepairPackageError("invalid-public-key", str(exc)) from exc

    try:
        contract = verify_signed_contract(raw_contract, {key_id: public_key_der})
    except (RepairContractSignatureError, RepairContractSchemaError) as exc:
        raise RepairPackageError(exc.code, str(exc)) from exc

    if original_path.name != contract.source_file_name:
        raise RepairPackageError(
            "filename-mismatch", f"original file name must be {contract.source_file_name!r}, got {original_path.name!r}"
        )
    # Not signed by the contract (see module docstring): a naming convention,
    # not a cryptographic guarantee that this is the file Lekta produced.
    if target_path.name != contract.output_policy.suggested_file_name:
        raise RepairPackageError(
            "filename-mismatch",
            f"target file name must be {contract.output_policy.suggested_file_name!r}, got {target_path.name!r}",
        )

    original_snapshot = capture_source(original_path)
    if original_snapshot.sha256 != contract.source_sha256:
        raise RepairPackageError("source-hash-mismatch")
    if original_snapshot.size != contract.source_size:
        raise RepairPackageError("source-size-mismatch")

    target_snapshot = capture_source(target_path)

    created_at = _parse_iso8601_utc(contract.created_at)
    expires_at = _parse_iso8601_utc(contract.expires_at)
    if now < created_at:
        raise RepairPackageError("invalid-time", "now is before createdAt")
    if now >= expires_at:
        raise RepairPackageError("expired")
    if (expires_at - created_at).total_seconds() > max_lifetime_seconds:
        raise RepairPackageError("lifetime-too-long")

    try:
        engine_ok = (
            _semver_tuple(contract.engine_min_version)
            <= _semver_tuple(engine_version)
            <= _semver_tuple(contract.engine_max_version)
        )
    except ValueError as exc:
        raise RepairPackageError("engine-out-of-range", str(exc)) from exc
    if not engine_ok:
        raise RepairPackageError("engine-out-of-range", engine_version)

    if output_dir.exists() and not output_dir.is_dir():
        raise RepairPackageError("invalid-output-dir", str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = reserve_output_path(
        output_dir, contract.output_policy.suggested_file_name, avoid={original_path, target_path}
    )

    contract_sha256 = sha256(contract_path.read_bytes()).hexdigest()

    return ValidatedRepairPackage(
        contract=contract,
        contract_sha256=contract_sha256,
        original_snapshot=original_snapshot,
        target_snapshot=target_snapshot,
        output_path=output_path,
    )
