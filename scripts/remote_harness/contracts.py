from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from hashlib import sha256
from typing import Any


class HarnessStatus(str, Enum):
    PASS = "PASS"
    WARN = "WARN"
    EXPECTED_BLOCK = "EXPECTED_BLOCK"
    ENGINE_FAIL = "ENGINE_FAIL"
    COM_FAIL = "COM_FAIL"
    TIMEOUT = "TIMEOUT"
    PROCESS_CRASH = "PROCESS_CRASH"
    SOURCE_MUTATED = "SOURCE_MUTATED"
    HARNESS_FAIL = "HARNESS_FAIL"


@dataclass(slots=True)
class StageResult:
    document: str
    stage: str
    status: HarnessStatus
    reason: str = ""
    elapsed_seconds: float = 0.0
    exit_code: int | None = None
    artifacts: dict[str, str] = field(default_factory=dict)
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        return data


@dataclass(slots=True)
class DocumentSummary:
    document: str
    source_sha256_before: str
    source_sha256_after: str = ""
    stages: dict[str, StageResult] = field(default_factory=dict)
    overall_status: HarnessStatus = HarnessStatus.PASS
    reason: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "document": self.document,
            "source_sha256_before": self.source_sha256_before,
            "source_sha256_after": self.source_sha256_after,
            "overall_status": self.overall_status.value,
            "reason": self.reason,
            "stages": {name: result.to_dict() for name, result in self.stages.items()},
        }


def normalize_failure(text: str) -> str:
    value = str(text or "")
    low = value.lower()
    if "-2147418111" in value or "call was rejected by callee" in low or "rpc_e_call_rejected" in low:
        return "RPC_E_CALL_REJECTED"
    if "0x800706ba" in low or "rpc server unavailable" in low or "rpc server is unavailable" in low:
        return "RPC_SERVER_UNAVAILABLE"
    if "'str' object has no attribute 'insertafter'" in low or "'str' object has no attribute 'paragraphformat'" in low:
        return "INVALID_ACTIVE_RANGE_TYPE"
    if "/word/header" in low and "error processing the xml" in low:
        return "INVALID_WORD_HEADER_XML"
    if "unsupported image format: .undefined" in low:
        return "IMAGE_FORMAT_UNDEFINED"
    if "image" in low and ("relationship" in low or "content type" in low or "corrupt" in low):
        return "IMAGE_RELATIONSHIP_ERROR"
    digest = sha256(value.encode("utf-8", errors="replace")).hexdigest()[:12]
    return f"UNCLASSIFIED:{digest}"
