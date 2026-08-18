"""Strict v1 parser for a Lekta Repair Contract (unsigned payload shape).

Mirrors the canonical schema in Lekta's docs/REPAIR_CONTRACT_V1.md and
src/repair/contract/contract-v1.ts. WordReplica does not execute fixer
requests (Lekta already did, producing the corrected target DOCX); this
parser only needs enough structural strictness to refuse a malformed or
tampered contract before Word ever opens. Deep per-fixer params validation
stays Lekta's responsibility.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from importlib import resources
from typing import Any

GOLDEN_GATES = ("G0", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9")
_UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.IGNORECASE
)
_SEMVER_RE = re.compile(r"^(0|[1-9]\d{0,3})\.(0|[1-9]\d{0,3})\.(0|[1-9]\d{0,3})$")
_ISO8601_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")
_REQUEST_ID_RE = re.compile(r"^req-\d{4}$")
_RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_UNSAFE_FILENAME_CHARS = set('<>:"/\\|?*')

with resources.files(__package__).joinpath("fixer_ids.json").open("r", encoding="utf-8") as _f:
    FIXER_IDS: frozenset[str] = frozenset(json.load(_f))

STANDALONE_DENIED_FIXER_IDS = frozenset({"footer-page-fixer"})

_TOP_LEVEL_KEYS = frozenset({
    "contractVersion", "jobId", "userId", "sourceSha256", "sourceSize", "sourceFileName",
    "createdAt", "expiresAt", "engineMinVersion", "engineMaxVersion", "requests",
    "allowedExceptions", "outputPolicy", "verificationPolicy", "contractSignature",
})
_REQUEST_KEYS = frozenset({"requestId", "fixerId", "ruleId", "params"})
_EXCEPTION_KEYS = frozenset({"requestId", "scope", "confirmationSha256", "confirmedAt"})
_OUTPUT_POLICY_KEYS = frozenset({"mode", "overwriteSource", "suggestedFileName"})
_VERIFICATION_POLICY_KEYS = frozenset({
    "requireSourceByteIdentity", "requireOpenAndRepairFalse", "requireVisibleTextEquality",
    "requireFieldsUpdateEquality", "preserveUnrelatedWordInstances", "requiredGates",
})
_EXCEPTION_SCOPES = frozenset({"metadata", "structure", "visible-text"})


class RepairContractSchemaError(ValueError):
    def __init__(self, code: str, path: str, message: str = "") -> None:
        super().__init__(f"{code} @ {path}: {message}" if message else f"{code} @ {path}")
        self.code = code
        self.path = path


def _fail(code: str, path: str, message: str = "") -> None:
    raise RepairContractSchemaError(code, path, message)


def _require_exact_keys(obj: Any, allowed: frozenset[str], path: str) -> None:
    if not isinstance(obj, dict):
        _fail("invalid-shape", path, "expected an object")
    extra = set(obj.keys()) - allowed
    missing = allowed - set(obj.keys())
    if extra or missing:
        _fail("invalid-shape", path, f"extra={sorted(extra)} missing={sorted(missing)}")


def _require_str(obj: dict, key: str, path: str, *, min_len: int = 1, max_len: int | None = None) -> str:
    value = obj.get(key)
    # bool is a subclass of int in Python, but never valid here regardless of key.
    if not isinstance(value, str) or isinstance(value, bool):
        _fail("invalid-shape", f"{path}/{key}", "expected a string")
    if len(value) < min_len or (max_len is not None and len(value) > max_len):
        _fail("invalid-shape", f"{path}/{key}", "length out of bounds")
    return value


def _require_bool(obj: dict, key: str, path: str) -> bool:
    value = obj.get(key)
    if not isinstance(value, bool):
        _fail("invalid-shape", f"{path}/{key}", "expected a boolean")
    return value


def _require_int(obj: dict, key: str, path: str, *, minimum: int = 0) -> int:
    value = obj.get(key)
    if isinstance(value, bool) or not isinstance(value, int):
        _fail("invalid-shape", f"{path}/{key}", "expected an integer")
    if value < minimum:
        _fail("invalid-shape", f"{path}/{key}", "out of bounds")
    return value


def _require_uuid(obj: dict, key: str, path: str) -> str:
    value = _require_str(obj, key, path)
    if not _UUID_RE.match(value):
        _fail("invalid-id", f"{path}/{key}", "not an RFC 4122 v1-5 UUID")
    return value


def _require_semver(obj: dict, key: str, path: str) -> str:
    value = _require_str(obj, key, path)
    if not _SEMVER_RE.match(value):
        _fail("invalid-shape", f"{path}/{key}", "not a valid SemVer")
    return value


def _require_iso8601(obj: dict, key: str, path: str) -> str:
    value = _require_str(obj, key, path)
    if not _ISO8601_RE.match(value):
        _fail("invalid-time", f"{path}/{key}", "not canonical UTC ISO-8601")
    return value


def _require_sha256_hex(obj: dict, key: str, path: str) -> str:
    value = _require_str(obj, key, path, min_len=64, max_len=64)
    if not re.fullmatch(r"[0-9a-f]{64}", value):
        _fail("invalid-hash", f"{path}/{key}", "not lowercase hex SHA-256")
    return value


def _require_docx_filename(obj: dict, key: str, path: str) -> str:
    value = _require_str(obj, key, path, min_len=1, max_len=180)
    if not value.lower().endswith(".docx"):
        _fail("invalid-shape", f"{path}/{key}", "must end with .docx")
    if "\x00" in value or ".." in value or "/" in value or "\\" in value:
        _fail("invalid-shape", f"{path}/{key}", "unsafe path characters")
    if any(character in _UNSAFE_FILENAME_CHARS for character in value):
        _fail("invalid-shape", f"{path}/{key}", "unsafe filename characters")
    stem = value.rsplit(".", 1)[0].upper()
    if stem in _RESERVED_WINDOWS_NAMES:
        _fail("invalid-shape", f"{path}/{key}", "reserved Windows device name")
    return value


@dataclass(frozen=True, slots=True)
class RepairContractRequestV1:
    request_id: str
    fixer_id: str
    rule_id: str
    params: dict[str, Any]


@dataclass(frozen=True, slots=True)
class AllowedExceptionV1:
    request_id: str
    scope: str
    confirmation_sha256: str
    confirmed_at: str


@dataclass(frozen=True, slots=True)
class OutputPolicyV1:
    mode: str
    overwrite_source: bool
    suggested_file_name: str


@dataclass(frozen=True, slots=True)
class VerificationPolicyV1:
    require_source_byte_identity: bool
    require_open_and_repair_false: bool
    require_visible_text_equality: bool
    require_fields_update_equality: bool
    preserve_unrelated_word_instances: bool
    required_gates: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class RepairContractSignatureV1:
    algorithm: str
    key_id: str
    value: str


@dataclass(frozen=True, slots=True)
class RepairContractV1:
    contract_version: int
    job_id: str
    user_id: str
    source_sha256: str
    source_size: int
    source_file_name: str
    created_at: str
    expires_at: str
    engine_min_version: str
    engine_max_version: str
    requests: tuple[RepairContractRequestV1, ...]
    allowed_exceptions: tuple[AllowedExceptionV1, ...]
    output_policy: OutputPolicyV1
    verification_policy: VerificationPolicyV1
    signature: RepairContractSignatureV1


def _parse_request(raw: dict, index: int) -> RepairContractRequestV1:
    path = f"requests/{index}"
    _require_exact_keys(raw, _REQUEST_KEYS, path)
    request_id = _require_str(raw, "requestId", path)
    if not _REQUEST_ID_RE.match(request_id):
        _fail("invalid-request-id", f"{path}/requestId", request_id)
    fixer_id = _require_str(raw, "fixerId", path)
    if fixer_id not in FIXER_IDS:
        _fail("unknown-fixer", f"{path}/fixerId", fixer_id)
    if fixer_id in STANDALONE_DENIED_FIXER_IDS:
        _fail("standalone-fixer-denied", f"{path}/fixerId", fixer_id)
    rule_id = _require_str(raw, "ruleId", path, min_len=1, max_len=200)
    params = raw.get("params")
    if not isinstance(params, dict):
        _fail("params-not-object", f"{path}/params")
    return RepairContractRequestV1(request_id=request_id, fixer_id=fixer_id, rule_id=rule_id, params=params)


def _parse_exception(raw: dict, index: int) -> AllowedExceptionV1:
    path = f"allowedExceptions/{index}"
    _require_exact_keys(raw, _EXCEPTION_KEYS, path)
    request_id = _require_str(raw, "requestId", path)
    scope = _require_str(raw, "scope", path)
    if scope not in _EXCEPTION_SCOPES:
        _fail("invalid-shape", f"{path}/scope", scope)
    confirmation_sha256 = _require_sha256_hex(raw, "confirmationSha256", path)
    confirmed_at = _require_iso8601(raw, "confirmedAt", path)
    return AllowedExceptionV1(
        request_id=request_id, scope=scope,
        confirmation_sha256=confirmation_sha256, confirmed_at=confirmed_at,
    )


def _parse_output_policy(raw: Any, path: str) -> OutputPolicyV1:
    _require_exact_keys(raw, _OUTPUT_POLICY_KEYS, path)
    mode = _require_str(raw, "mode", path)
    if mode != "new-file":
        _fail("unsafe-output-policy", f"{path}/mode", mode)
    overwrite_source = _require_bool(raw, "overwriteSource", path)
    if overwrite_source is not False:
        _fail("unsafe-output-policy", f"{path}/overwriteSource", "must be false")
    suggested_file_name = _require_docx_filename(raw, "suggestedFileName", path)
    return OutputPolicyV1(mode=mode, overwrite_source=overwrite_source, suggested_file_name=suggested_file_name)


def _parse_verification_policy(raw: Any, path: str) -> VerificationPolicyV1:
    _require_exact_keys(raw, _VERIFICATION_POLICY_KEYS, path)
    flags = {
        key: _require_bool(raw, key, path)
        for key in (
            "requireSourceByteIdentity", "requireOpenAndRepairFalse", "requireVisibleTextEquality",
            "requireFieldsUpdateEquality", "preserveUnrelatedWordInstances",
        )
    }
    if not all(flags.values()):
        _fail("insufficient-verification-policy", path, str(flags))
    gates = raw.get("requiredGates")
    if not isinstance(gates, list) or tuple(gates) != GOLDEN_GATES:
        _fail("insufficient-verification-policy", f"{path}/requiredGates", str(gates))
    return VerificationPolicyV1(
        require_source_byte_identity=flags["requireSourceByteIdentity"],
        require_open_and_repair_false=flags["requireOpenAndRepairFalse"],
        require_visible_text_equality=flags["requireVisibleTextEquality"],
        require_fields_update_equality=flags["requireFieldsUpdateEquality"],
        preserve_unrelated_word_instances=flags["preserveUnrelatedWordInstances"],
        required_gates=tuple(gates),
    )


def _parse_signature(raw: Any, path: str) -> RepairContractSignatureV1:
    _require_exact_keys(raw, frozenset({"algorithm", "keyId", "value"}), path)
    algorithm = _require_str(raw, "algorithm", path)
    key_id = _require_str(raw, "keyId", path, min_len=1, max_len=80)
    if not re.fullmatch(r"[A-Za-z0-9._-]{1,80}", key_id):
        _fail("invalid-key-id", f"{path}/keyId", key_id)
    value = _require_str(raw, "value", path)
    return RepairContractSignatureV1(algorithm=algorithm, key_id=key_id, value=value)


def parse_repair_contract_v1(raw: Any) -> RepairContractV1:
    """Strictly parse an already-canonical (or plain) unsigned/signed payload dict.

    Does not verify the signature; callers that need a trust boundary must go
    through repair_contract.signature.verify_signed_contract instead.
    """
    _require_exact_keys(raw, _TOP_LEVEL_KEYS, "$")
    version = raw.get("contractVersion")
    if isinstance(version, bool) or version != 1:
        _fail("unsupported-version", "$/contractVersion", str(version))
    job_id = _require_uuid(raw, "jobId", "$")
    user_id = _require_uuid(raw, "userId", "$")
    source_sha256 = _require_sha256_hex(raw, "sourceSha256", "$")
    source_size = _require_int(raw, "sourceSize", "$", minimum=0)
    source_file_name = _require_docx_filename(raw, "sourceFileName", "$")
    created_at = _require_iso8601(raw, "createdAt", "$")
    expires_at = _require_iso8601(raw, "expiresAt", "$")
    if expires_at <= created_at:
        _fail("invalid-time", "$/expiresAt", "must be strictly after createdAt")
    engine_min = _require_semver(raw, "engineMinVersion", "$")
    engine_max = _require_semver(raw, "engineMaxVersion", "$")
    if _semver_tuple(engine_max) < _semver_tuple(engine_min):
        _fail("invalid-shape", "$/engineMaxVersion", "must be >= engineMinVersion")

    requests_raw = raw.get("requests")
    if not isinstance(requests_raw, list) or not (1 <= len(requests_raw) <= 64):
        _fail("request-count", "$/requests", str(requests_raw))
    requests = tuple(_parse_request(item, index) for index, item in enumerate(requests_raw))
    seen_ids = set()
    for index, request in enumerate(requests):
        if request.request_id in seen_ids:
            _fail("duplicate-request-id", f"requests/{index}/requestId", request.request_id)
        seen_ids.add(request.request_id)

    exceptions_raw = raw.get("allowedExceptions")
    if not isinstance(exceptions_raw, list) or len(exceptions_raw) > 64:
        _fail("invalid-shape", "$/allowedExceptions", str(exceptions_raw))
    exceptions = tuple(_parse_exception(item, index) for index, item in enumerate(exceptions_raw))
    exception_request_ids = {exception.request_id for exception in exceptions}
    request_ids = {request.request_id for request in requests}
    for exception in exceptions:
        if exception.request_id not in request_ids:
            _fail("orphan-exception", "$/allowedExceptions", exception.request_id)

    output_policy = _parse_output_policy(raw.get("outputPolicy"), "$/outputPolicy")
    verification_policy = _parse_verification_policy(raw.get("verificationPolicy"), "$/verificationPolicy")
    signature = _parse_signature(raw.get("contractSignature"), "$/contractSignature")

    return RepairContractV1(
        contract_version=1,
        job_id=job_id,
        user_id=user_id,
        source_sha256=source_sha256,
        source_size=source_size,
        source_file_name=source_file_name,
        created_at=created_at,
        expires_at=expires_at,
        engine_min_version=engine_min,
        engine_max_version=engine_max,
        requests=requests,
        allowed_exceptions=exceptions,
        output_policy=output_policy,
        verification_policy=verification_policy,
        signature=signature,
    )


def _semver_tuple(value: str) -> tuple[int, int, int]:
    major, minor, patch = value.split(".")
    return int(major), int(minor), int(patch)
