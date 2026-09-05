"""DPAPI-protected state for resuming one claimed Lekta repair job.

The file contains only a schema marker, job id and DPAPI ciphertext. The
device private key, document/contract hashes, checkpoint and completion
receipt stay inside the current-Windows-user protected payload. A sibling
per-job package workspace exists only while reconstruction is resumable and is
removed together with the encrypted state after terminal completion.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol
from urllib.parse import urlsplit

from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.serialization import (
    Encoding,
    NoEncryption,
    PrivateFormat,
    PublicFormat,
    load_der_private_key,
)
from platformdirs import user_data_dir

from word_replica.runner.lekta_claim import DeviceIdentity
from word_replica.runner.lekta_status import (
    StatusProtocolError,
    validate_signed_local_repair_status,
)


SCHEMA_VERSION = 1
MAX_STATE_BYTES = 128 * 1024
_UUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_ENGINE_VERSION = re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+(?:[-+][A-Za-z0-9.-]+)?$")
_PHASES = {
    "claimed", "processing", "retryable", "completed_pending_receipt",
    "failed_pending_receipt",
}
_INNER_KEYS = {
    "schemaVersion", "jobId", "sourceSha256", "contractSha256",
    "engineVersion", "contractVersion", "phase", "sequence",
    "checkpointSha256", "outputPath", "outputSha256", "reportSha256",
    "pendingStatus", "devicePrivateKeyPkcs8", "devicePublicKeySpki",
    "claimContractJson", "sourceDownloadUrl", "targetDownloadUrl",
}
_OUTER_KEYS = {"version", "jobId", "protectedData"}
_CLAIM_ATTEMPT_KEYS = {
    "schemaVersion", "jobId", "claimToken",
    "devicePrivateKeyPkcs8", "devicePublicKeySpki",
}
_CLAIM_ENTROPY_PREFIX = b"WordReplica/LektaClaimAttempt/v1\x00"

_ENTROPY_PREFIX = b"WordReplica/LektaRetry/v1\x00"


class RetryStateError(RuntimeError):
    """Encrypted retry state is unavailable, invalid or bound elsewhere."""


class DataProtector(Protocol):
    def protect(self, data: bytes, *, entropy: bytes) -> bytes: ...
    def unprotect(self, data: bytes, *, entropy: bytes) -> bytes: ...


def _base64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _decode_base64url(value: Any) -> bytes:
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9_-]*", value) or len(value) % 4 == 1:
        raise ValueError("invalid base64url")
    data = base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))
    if _base64url(data) != value:
        raise ValueError("non-canonical base64url")
    return data


class WindowsDpapiProtector:
    def protect(self, data: bytes, *, entropy: bytes) -> bytes:
        try:
            import win32crypt
            import win32cryptcon

            return win32crypt.CryptProtectData(
                data, "WordReplica Lekta retry", entropy, None, None,
                win32cryptcon.CRYPTPROTECT_UI_FORBIDDEN,
            )
        except Exception as exc:
            raise RetryStateError("DPAPI protect failed") from exc

    def unprotect(self, data: bytes, *, entropy: bytes) -> bytes:
        try:
            import win32crypt
            import win32cryptcon

            _description, plaintext = win32crypt.CryptUnprotectData(
                data, entropy, None, None, win32cryptcon.CRYPTPROTECT_UI_FORBIDDEN,
            )
            return plaintext
        except Exception as exc:
            raise RetryStateError("DPAPI unprotect failed") from exc


def _optional_hash(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not _SHA256.fullmatch(value):
        raise ValueError("invalid SHA-256")
    return value


def _optional_https_url(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError("invalid claim download URL")
    parsed = urlsplit(value)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise ValueError("invalid claim download URL")
    return value


@dataclass(frozen=True, slots=True)
class ClaimAttemptState:
    schema_version: int
    job_id: str
    claim_token: str = field(repr=False)
    device_private_key_pkcs8: bytes = field(repr=False)
    device_public_key_spki: str

    @classmethod
    def build(
        cls, *, job_id: str, claim_token: str, identity: DeviceIdentity
    ) -> "ClaimAttemptState":
        state = cls(
            schema_version=SCHEMA_VERSION,
            job_id=job_id,
            claim_token=claim_token,
            device_private_key_pkcs8=identity.private_key.private_bytes(
                Encoding.DER, PrivateFormat.PKCS8, NoEncryption()
            ),
            device_public_key_spki=identity.public_key_spki,
        )
        state._validate()
        return state

    def _validate(self) -> None:
        if self.schema_version != SCHEMA_VERSION:
            raise ValueError("invalid claim attempt schema")
        if not isinstance(self.job_id, str) or not _UUID.fullmatch(self.job_id):
            raise ValueError("invalid claim attempt job")
        try:
            token_bytes = _decode_base64url(self.claim_token)
        except ValueError as exc:
            raise ValueError("invalid claim attempt token") from exc
        if len(self.claim_token) != 43 or len(token_bytes) != 32:
            raise ValueError("invalid claim attempt token")
        private_key = load_der_private_key(self.device_private_key_pkcs8, password=None)
        if not isinstance(private_key, ec.EllipticCurvePrivateKey) or not isinstance(
            private_key.curve, ec.SECP256R1
        ):
            raise ValueError("invalid claim attempt device key")
        public_spki = private_key.public_key().public_bytes(
            Encoding.DER, PublicFormat.SubjectPublicKeyInfo
        )
        if _base64url(public_spki) != self.device_public_key_spki:
            raise ValueError("claim attempt device key mismatch")

    @property
    def identity(self) -> DeviceIdentity:
        private_key = load_der_private_key(self.device_private_key_pkcs8, password=None)
        if not isinstance(private_key, ec.EllipticCurvePrivateKey):
            raise RetryStateError("invalid encrypted claim attempt")
        return DeviceIdentity(
            private_key=private_key,
            public_key_spki=self.device_public_key_spki,
        )

    def to_json(self) -> dict:
        self._validate()
        return {
            "schemaVersion": self.schema_version,
            "jobId": self.job_id,
            "claimToken": self.claim_token,
            "devicePrivateKeyPkcs8": _base64url(self.device_private_key_pkcs8),
            "devicePublicKeySpki": self.device_public_key_spki,
        }

    @classmethod
    def from_json(cls, value: Any) -> "ClaimAttemptState":
        if not isinstance(value, dict) or set(value) != _CLAIM_ATTEMPT_KEYS:
            raise ValueError("invalid claim attempt shape")
        state = cls(
            schema_version=value["schemaVersion"],
            job_id=value["jobId"],
            claim_token=value["claimToken"],
            device_private_key_pkcs8=_decode_base64url(value["devicePrivateKeyPkcs8"]),
            device_public_key_spki=value["devicePublicKeySpki"],
        )
        state._validate()
        return state


@dataclass(frozen=True, slots=True)
class RetryState:
    schema_version: int
    job_id: str
    source_sha256: str
    contract_sha256: str
    engine_version: str
    contract_version: int
    phase: str
    sequence: int
    checkpoint_sha256: str | None
    output_path: str | None
    output_sha256: str | None
    report_sha256: str | None
    pending_status: dict | None
    claim_contract_json: str | None
    source_download_url: str | None
    target_download_url: str | None
    device_private_key_pkcs8: bytes = field(repr=False)
    device_public_key_spki: str

    @classmethod
    def build(
        cls,
        *,
        job_id: str,
        source_sha256: str,
        contract_sha256: str,
        engine_version: str,
        contract_version: int,
        phase: str,
        sequence: int,
        checkpoint_sha256: str | None,
        output_path: str | None,
        output_sha256: str | None,
        report_sha256: str | None,
        pending_status: dict | None,
        claim_contract_json: str | None = None,
        source_download_url: str | None = None,
        target_download_url: str | None = None,
        identity: DeviceIdentity,
    ) -> "RetryState":
        private_der = identity.private_key.private_bytes(
            Encoding.DER, PrivateFormat.PKCS8, NoEncryption()
        )
        state = cls(
            schema_version=SCHEMA_VERSION,
            job_id=job_id,
            source_sha256=source_sha256,
            contract_sha256=contract_sha256,
            engine_version=engine_version,
            contract_version=contract_version,
            phase=phase,
            sequence=sequence,
            checkpoint_sha256=checkpoint_sha256,
            output_path=output_path,
            output_sha256=output_sha256,
            report_sha256=report_sha256,
            pending_status=pending_status,
            claim_contract_json=claim_contract_json,
            source_download_url=source_download_url,
            target_download_url=target_download_url,
            device_private_key_pkcs8=private_der,
            device_public_key_spki=identity.public_key_spki,
        )
        state._validate()
        return state

    def _validate(self) -> None:
        if self.schema_version != SCHEMA_VERSION:
            raise ValueError("invalid retry schema")
        if not isinstance(self.job_id, str) or not _UUID.fullmatch(self.job_id):
            raise ValueError("invalid retry job")
        if not isinstance(self.source_sha256, str) or not _SHA256.fullmatch(self.source_sha256):
            raise ValueError("invalid source SHA-256")
        if not isinstance(self.contract_sha256, str) or not _SHA256.fullmatch(self.contract_sha256):
            raise ValueError("invalid contract SHA-256")
        if not isinstance(self.engine_version, str) or not _ENGINE_VERSION.fullmatch(self.engine_version):
            raise ValueError("invalid engine version")
        if self.contract_version != 1 or self.phase not in _PHASES:
            raise ValueError("invalid retry binding")
        if isinstance(self.sequence, bool) or not isinstance(self.sequence, int) or self.sequence < 0:
            raise ValueError("invalid retry sequence")
        _optional_hash(self.checkpoint_sha256)
        _optional_hash(self.output_sha256)
        _optional_hash(self.report_sha256)
        _optional_https_url(self.source_download_url)
        _optional_https_url(self.target_download_url)
        if self.output_path is not None and (
            not isinstance(self.output_path, str) or not Path(self.output_path).is_absolute()
        ):
            raise ValueError("invalid retry output path")
        if self.phase == "claimed" and (
            self.sequence != 0
            or any((self.checkpoint_sha256, self.output_path, self.output_sha256, self.report_sha256, self.pending_status))
            or not isinstance(self.claim_contract_json, str)
            or self.source_download_url is None
            or self.target_download_url is None
        ):
            raise ValueError("invalid claimed retry state")
        if self.phase == "claimed":
            try:
                contract_bytes = self.claim_contract_json.encode("utf-8")
                raw_contract = json.loads(self.claim_contract_json)
            except (UnicodeError, ValueError, AttributeError) as exc:
                raise ValueError("invalid claimed retry state") from exc
            if (
                not isinstance(raw_contract, dict)
                or hashlib.sha256(contract_bytes).hexdigest() != self.contract_sha256
            ):
                raise ValueError("invalid claimed retry state")
        elif any((self.claim_contract_json, self.source_download_url, self.target_download_url)):
            raise ValueError("claim data persisted beyond claimed phase")
        if self.phase == "processing" and (
            self.sequence < 1 or any((self.checkpoint_sha256, self.output_path, self.output_sha256, self.report_sha256))
        ):
            raise ValueError("invalid processing retry state")
        if self.phase == "retryable" and (
            self.sequence < 1 or self.checkpoint_sha256 is None
            or any((self.output_path, self.output_sha256, self.report_sha256))
        ):
            raise ValueError("invalid retryable state")
        if self.phase == "completed_pending_receipt" and (
            self.sequence < 2 or self.output_path is None
            or self.output_sha256 is None or self.report_sha256 is None
            or self.pending_status is None
        ):
            raise ValueError("invalid completion retry state")
        if self.phase == "failed_pending_receipt" and (
            self.sequence < 1 or any((self.checkpoint_sha256, self.output_path, self.output_sha256, self.report_sha256))
            or self.pending_status is None
        ):
            raise ValueError("invalid terminal failure retry state")
        if self.pending_status is not None:
            try:
                pending_job, pending_sequence, pending_event = validate_signed_local_repair_status(
                    self.pending_status
                )
            except StatusProtocolError as exc:
                raise ValueError("invalid pending status") from exc
            expected_events = {
                "processing": {"processing", "heartbeat"},
                "retryable": {"retryable"},
                "completed_pending_receipt": {"completed"},
                "failed_pending_receipt": {"local_failed"},
            }.get(self.phase, set())
            if pending_job != self.job_id or pending_sequence != self.sequence or pending_event not in expected_events:
                raise ValueError("pending status binding mismatch")
            if pending_event == "retryable" and self.pending_status["checkpointSha256"] != self.checkpoint_sha256:
                raise ValueError("pending status binding mismatch")
            if pending_event == "completed" and (
                self.pending_status["outputSha256"] != self.output_sha256
                or self.pending_status["reportSha256"] != self.report_sha256
            ):
                raise ValueError("pending status binding mismatch")
        private_key = load_der_private_key(self.device_private_key_pkcs8, password=None)
        if not isinstance(private_key, ec.EllipticCurvePrivateKey) or not isinstance(private_key.curve, ec.SECP256R1):
            raise ValueError("invalid retry device key")
        public_spki = private_key.public_key().public_bytes(Encoding.DER, PublicFormat.SubjectPublicKeyInfo)
        if _base64url(public_spki) != self.device_public_key_spki:
            raise ValueError("retry device key mismatch")

    @property
    def identity(self) -> DeviceIdentity:
        private_key = load_der_private_key(self.device_private_key_pkcs8, password=None)
        if not isinstance(private_key, ec.EllipticCurvePrivateKey):
            raise RetryStateError("invalid encrypted retry state")
        return DeviceIdentity(private_key=private_key, public_key_spki=self.device_public_key_spki)

    def to_json(self) -> dict:
        self._validate()
        return {
            "schemaVersion": self.schema_version,
            "jobId": self.job_id,
            "sourceSha256": self.source_sha256,
            "contractSha256": self.contract_sha256,
            "engineVersion": self.engine_version,
            "contractVersion": self.contract_version,
            "phase": self.phase,
            "sequence": self.sequence,
            "checkpointSha256": self.checkpoint_sha256,
            "outputPath": self.output_path,
            "outputSha256": self.output_sha256,
            "reportSha256": self.report_sha256,
            "pendingStatus": self.pending_status,
            "claimContractJson": self.claim_contract_json,
            "sourceDownloadUrl": self.source_download_url,
            "targetDownloadUrl": self.target_download_url,
            "devicePrivateKeyPkcs8": _base64url(self.device_private_key_pkcs8),
            "devicePublicKeySpki": self.device_public_key_spki,
        }

    @classmethod
    def from_json(cls, value: Any) -> "RetryState":
        if not isinstance(value, dict) or set(value) != _INNER_KEYS:
            raise ValueError("invalid retry state shape")
        state = cls(
            schema_version=value["schemaVersion"],
            job_id=value["jobId"],
            source_sha256=value["sourceSha256"],
            contract_sha256=value["contractSha256"],
            engine_version=value["engineVersion"],
            contract_version=value["contractVersion"],
            phase=value["phase"],
            sequence=value["sequence"],
            checkpoint_sha256=value["checkpointSha256"],
            output_path=value["outputPath"],
            output_sha256=value["outputSha256"],
            report_sha256=value["reportSha256"],
            pending_status=value["pendingStatus"],
            claim_contract_json=value["claimContractJson"],
            source_download_url=value["sourceDownloadUrl"],
            target_download_url=value["targetDownloadUrl"],
            device_private_key_pkcs8=_decode_base64url(value["devicePrivateKeyPkcs8"]),
            device_public_key_spki=value["devicePublicKeySpki"],
        )
        state._validate()
        return state


class SecureRetryStore:
    def __init__(self, root: Path | None = None, *, protector: DataProtector | None = None) -> None:
        self.root = Path(root) if root is not None else (
            Path(user_data_dir("WordReplica", "WordReplica")) / "lekta_runner" / "active_jobs"
        )
        self.protector = protector or WindowsDpapiProtector()

    @staticmethod
    def _validate_job_id(job_id: str) -> str:
        if not isinstance(job_id, str) or not _UUID.fullmatch(job_id):
            raise RetryStateError("invalid retry job")
        return job_id

    def path_for(self, job_id: str) -> Path:
        return self.root / self._validate_job_id(job_id) / "retry-state.json"

    def claim_attempt_path_for(self, job_id: str) -> Path:
        return self.root / self._validate_job_id(job_id) / "claim-attempt.json"

    def workspace_for(self, job_id: str) -> Path:
        return self.root / self._validate_job_id(job_id) / "package"

    @staticmethod
    def _entropy(job_id: str) -> bytes:
        return _ENTROPY_PREFIX + job_id.encode("ascii")

    @staticmethod
    def _claim_entropy(job_id: str) -> bytes:
        return _CLAIM_ENTROPY_PREFIX + job_id.encode("ascii")

    def save_claim_attempt(
        self, *, job_id: str, claim_token: str, identity: DeviceIdentity
    ) -> Path:
        try:
            state = ClaimAttemptState.build(
                job_id=job_id, claim_token=claim_token, identity=identity
            )
        except (TypeError, ValueError) as exc:
            raise RetryStateError("invalid claim attempt state") from exc
        path = self.claim_attempt_path_for(job_id)
        if path.exists():
            existing = self.load_claim_attempt(job_id)
            if existing.claim_token != state.claim_token:
                raise RetryStateError("claim attempt binding mismatch")
            return path

        plaintext = json.dumps(
            state.to_json(), ensure_ascii=False, separators=(",", ":"), sort_keys=True,
        ).encode("utf-8")
        protected = self.protector.protect(
            plaintext, entropy=self._claim_entropy(job_id)
        )
        envelope = json.dumps({
            "version": SCHEMA_VERSION,
            "jobId": job_id,
            "protectedData": _base64url(protected),
        }, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
        if len(envelope) > MAX_STATE_BYTES:
            raise RetryStateError("encrypted claim attempt is too large")

        path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix="claim-attempt-", suffix=".tmp", dir=path.parent
        )
        temporary = Path(temporary_name)
        try:
            with os.fdopen(descriptor, "wb") as target:
                target.write(envelope)
                target.flush()
                os.fsync(target.fileno())
            try:
                # Atomic first-writer-wins publication: link only succeeds when
                # no competing process has already bound this job.
                os.link(temporary, path)
            except FileExistsError:
                existing = self.load_claim_attempt(job_id)
                if existing.claim_token != state.claim_token:
                    raise RetryStateError("claim attempt binding mismatch")
        finally:
            temporary.unlink(missing_ok=True)
        return path

    def load_claim_attempt(self, job_id: str) -> ClaimAttemptState:
        path = self.claim_attempt_path_for(job_id)
        try:
            if not path.is_file() or path.stat().st_size > MAX_STATE_BYTES:
                raise ValueError("missing or oversized claim attempt")
            envelope = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(envelope, dict) or set(envelope) != _OUTER_KEYS:
                raise ValueError("invalid claim attempt envelope")
            if envelope["version"] != SCHEMA_VERSION or envelope["jobId"] != job_id:
                raise ValueError("claim attempt envelope mismatch")
            protected = _decode_base64url(envelope["protectedData"])
            plaintext = self.protector.unprotect(
                protected, entropy=self._claim_entropy(job_id)
            )
            state = ClaimAttemptState.from_json(json.loads(plaintext.decode("utf-8")))
            if state.job_id != job_id:
                raise ValueError("claim attempt payload mismatch")
            return state
        except RetryStateError as exc:
            raise RetryStateError("invalid encrypted claim attempt") from exc
        except (OSError, UnicodeError, ValueError, TypeError) as exc:
            raise RetryStateError("invalid encrypted claim attempt") from exc

    def delete_claim_attempt(self, job_id: str) -> None:
        self.claim_attempt_path_for(job_id).unlink(missing_ok=True)

    def save(self, state: RetryState) -> Path:
        try:
            state._validate()
        except (TypeError, ValueError) as exc:
            raise RetryStateError("invalid retry state") from exc
        path = self.path_for(state.job_id)
        if path.exists():
            existing = self.load(state.job_id)
            immutable_existing = (
                existing.job_id, existing.source_sha256, existing.contract_sha256,
                existing.engine_version, existing.contract_version, existing.device_public_key_spki,
            )
            immutable_new = (
                state.job_id, state.source_sha256, state.contract_sha256,
                state.engine_version, state.contract_version, state.device_public_key_spki,
            )
            if immutable_existing != immutable_new:
                raise RetryStateError("retry binding mismatch")
            if state.sequence < existing.sequence:
                raise RetryStateError("retry sequence rollback")

        plaintext = json.dumps(
            state.to_json(), ensure_ascii=False, separators=(",", ":"), sort_keys=True,
        ).encode("utf-8")
        protected = self.protector.protect(plaintext, entropy=self._entropy(state.job_id))
        envelope = json.dumps({
            "version": SCHEMA_VERSION,
            "jobId": state.job_id,
            "protectedData": _base64url(protected),
        }, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
        if len(envelope) > MAX_STATE_BYTES:
            raise RetryStateError("encrypted retry state is too large")

        path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix="retry-state-", suffix=".tmp", dir=path.parent
        )
        temporary = Path(temporary_name)
        try:
            with os.fdopen(descriptor, "wb") as target:
                target.write(envelope)
                target.flush()
                os.fsync(target.fileno())
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)
        return path

    def load(self, job_id: str) -> RetryState:
        path = self.path_for(job_id)
        try:
            if not path.is_file() or path.stat().st_size > MAX_STATE_BYTES:
                raise ValueError("missing or oversized state")
            envelope = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(envelope, dict) or set(envelope) != _OUTER_KEYS:
                raise ValueError("invalid envelope")
            if envelope["version"] != SCHEMA_VERSION or envelope["jobId"] != job_id:
                raise ValueError("retry envelope mismatch")
            protected = _decode_base64url(envelope["protectedData"])
            plaintext = self.protector.unprotect(protected, entropy=self._entropy(job_id))
            state = RetryState.from_json(json.loads(plaintext.decode("utf-8")))
            if state.job_id != job_id:
                raise ValueError("retry payload mismatch")
            return state
        except RetryStateError as exc:
            raise RetryStateError("invalid encrypted retry state") from exc
        except (OSError, UnicodeError, ValueError, TypeError) as exc:
            raise RetryStateError("invalid encrypted retry state") from exc

    def delete(self, job_id: str) -> None:
        job_dir = self.path_for(job_id).parent
        if job_dir.parent.resolve() != self.root.resolve():
            raise RetryStateError("invalid retry cleanup target")
        if job_dir.exists():
            shutil.rmtree(job_dir)
