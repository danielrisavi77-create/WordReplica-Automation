"""Machine-readable completion report for one repair-package run.

`open_and_repair` and `fields_update_equal` are populated by real Word-COM
probes (word_replica.qa.word_render.detect_open_and_repair and
check_fields_update_equality, wired as RepairPackageService's defaults).
Both only need Word to be installed, not activated — live-verified against
this machine's unactivated Word. If Word is unavailable at all, or a caller
passes its own checker that can't determine an answer, the field stays
`None` ("not verified"); `None` counts as failing (not passing) full_pass,
the same fail-closed default the rest of this package uses for anything
unverified.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

SCHEMA_VERSION = 1
STATUS_FULL_PASS = "FULL_PASS"
STATUS_RETRYABLE = "RETRYABLE"
STATUS_FAILED = "FAILED"


@dataclass(frozen=True, slots=True)
class RepairCompletionReport:
    schema_version: int
    job_id: str
    status: str
    contract_sha256: str
    source_sha256: str
    source_size: int
    target_sha256: str
    target_size: int
    output_sha256: str | None
    output_size: int | None
    output_path: str | None
    project_id: str | None
    blueprint_fingerprint: str | None
    reconstruction_status: str | None
    original_unchanged: bool
    open_and_repair: bool | None
    visible_text_equal: bool | None
    fields_update_equal: bool | None
    gates: dict[str, bool]
    full_pass: bool
    reasons: tuple[str, ...]
    timing_seconds: dict[str, float]
    word_ownership_evidence_path: str | None

    def to_json(self) -> dict[str, Any]:
        data = asdict(self)
        data["reasons"] = list(self.reasons)
        return data


def compute_full_pass(
    *,
    reconstruction_status: str | None,
    original_unchanged: bool,
    gates: dict[str, bool],
    open_and_repair: bool | None,
    visible_text_equal: bool | None,
    fields_update_equal: bool | None,
    required_gates: tuple[str, ...],
) -> bool:
    if reconstruction_status != "PASS":
        return False
    if not original_unchanged:
        return False
    if open_and_repair is not False:  # None (unverified) or True both fail closed.
        return False
    if visible_text_equal is not True:
        return False
    if fields_update_equal is not True:
        return False
    return all(gates.get(name) is True for name in required_gates)


def build_completion_report(
    *,
    job_id: str,
    contract_sha256: str,
    source_sha256: str,
    source_size: int,
    target_sha256: str,
    target_size: int,
    required_gates: tuple[str, ...],
    output_sha256: str | None = None,
    output_size: int | None = None,
    output_path: str | None = None,
    project_id: str | None = None,
    blueprint_fingerprint: str | None = None,
    reconstruction_status: str | None = None,
    original_unchanged: bool = False,
    open_and_repair: bool | None = None,
    visible_text_equal: bool | None = None,
    fields_update_equal: bool | None = None,
    gates: dict[str, bool] | None = None,
    reasons: tuple[str, ...] = (),
    timing_seconds: dict[str, float] | None = None,
    word_ownership_evidence_path: str | None = None,
) -> RepairCompletionReport:
    gates = dict(gates or {})
    full_pass = compute_full_pass(
        reconstruction_status=reconstruction_status,
        original_unchanged=original_unchanged,
        gates=gates,
        open_and_repair=open_and_repair,
        visible_text_equal=visible_text_equal,
        fields_update_equal=fields_update_equal,
        required_gates=required_gates,
    )
    if full_pass:
        status = STATUS_FULL_PASS
    elif reconstruction_status == "FAIL" or reconstruction_status is None:
        status = STATUS_FAILED
    else:
        status = STATUS_RETRYABLE
    return RepairCompletionReport(
        schema_version=SCHEMA_VERSION,
        job_id=job_id,
        status=status,
        contract_sha256=contract_sha256,
        source_sha256=source_sha256,
        source_size=source_size,
        target_sha256=target_sha256,
        target_size=target_size,
        output_sha256=output_sha256,
        output_size=output_size,
        output_path=output_path,
        project_id=project_id,
        blueprint_fingerprint=blueprint_fingerprint,
        reconstruction_status=reconstruction_status,
        original_unchanged=original_unchanged,
        open_and_repair=open_and_repair,
        visible_text_equal=visible_text_equal,
        fields_update_equal=fields_update_equal,
        gates=gates,
        full_pass=full_pass,
        reasons=reasons,
        timing_seconds=dict(timing_seconds or {}),
        word_ownership_evidence_path=word_ownership_evidence_path,
    )
