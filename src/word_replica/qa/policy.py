from dataclasses import dataclass, field
from hashlib import sha256
from typing import Any

from word_replica.domain.enums import RunStatus
from word_replica.domain.results import WarningItem
from word_replica.qa.content import l0_projection
from word_replica.qa.formatting import l2_projection
from word_replica.qa.layout import l3_projection
from word_replica.qa.structure import l1_projection


@dataclass(slots=True)
class QaFinding:
    code: str
    path: str
    expected: object
    actual: object
    severity: str


@dataclass(slots=True)
class QaLevelResult:
    level: str
    passed: bool
    findings: list[QaFinding | str] = field(default_factory=list)
    has_required_loss: bool = False


@dataclass(slots=True)
class QaBundle:
    levels: dict[str, QaLevelResult]
    render: Any | None = None


def _safe_value(value: Any) -> Any:
    if isinstance(value, str) and len(value) > 200:
        return {
            "excerpt": value[:200] + "…",
            "sha256": sha256(value.encode("utf-8")).hexdigest(),
            "length": len(value),
        }
    return value


def compare_projection(level: str, expected: Any, actual: Any) -> list[QaFinding]:
    findings: list[QaFinding] = []
    severity = "error" if level in {"L0", "L1"} else "warning"

    def walk(left: Any, right: Any, path: str) -> None:
        if type(left) is not type(right):
            findings.append(
                QaFinding(
                    f"{level}_MISMATCH",
                    path,
                    _safe_value(left),
                    _safe_value(right),
                    severity,
                )
            )
            return
        if isinstance(left, dict):
            keys = sorted(set(left) | set(right), key=str)
            for key in keys:
                child_path = f"{path}/{key}" if path else str(key)
                if key not in left or key not in right:
                    findings.append(
                        QaFinding(
                            f"{level}_MISMATCH",
                            child_path,
                            _safe_value(left.get(key, "<missing>")),
                            _safe_value(right.get(key, "<missing>")),
                            severity,
                        )
                    )
                else:
                    walk(left[key], right[key], child_path)
            return
        if isinstance(left, (list, tuple)):
            max_len = max(len(left), len(right))
            for index in range(max_len):
                child_path = f"{path}/{index}" if path else str(index)
                if index >= len(left) or index >= len(right):
                    findings.append(
                        QaFinding(
                            f"{level}_MISMATCH",
                            child_path,
                            _safe_value(left[index] if index < len(left) else "<missing>"),
                            _safe_value(right[index] if index < len(right) else "<missing>"),
                            severity,
                        )
                    )
                else:
                    walk(left[index], right[index], child_path)
            return
        if left != right:
            findings.append(
                QaFinding(
                    f"{level}_MISMATCH",
                    path or "/",
                    _safe_value(left),
                    _safe_value(right),
                    severity,
                )
            )

    walk(expected, actual, "")
    return findings


def run_l0_l3(source_model, rebuilt_model) -> QaBundle:
    projections = {
        "L0": (l0_projection(source_model), l0_projection(rebuilt_model)),
        "L1": (l1_projection(source_model), l1_projection(rebuilt_model)),
        "L2": (l2_projection(source_model), l2_projection(rebuilt_model)),
        "L3": (l3_projection(source_model), l3_projection(rebuilt_model)),
    }
    levels: dict[str, QaLevelResult] = {}
    for level, (expected, actual) in projections.items():
        findings = compare_projection(level, expected, actual)
        levels[level] = QaLevelResult(
            level=level,
            passed=not findings,
            findings=findings,
            has_required_loss=bool(findings) if level == "L1" else False,
        )
    return QaBundle(levels=levels)


def classify_run(
    bundle: QaBundle,
    warnings: list[WarningItem],
    critical_reasons: list[str] | None = None,
) -> RunStatus:
    if critical_reasons:
        return RunStatus.FAIL
    l0 = bundle.levels.get("L0")
    if l0 is not None and not l0.passed:
        return RunStatus.FAIL
    l1 = bundle.levels.get("L1")
    if l1 is not None and not l1.passed and l1.has_required_loss:
        return RunStatus.FAIL
    if any(warning.affects_status for warning in warnings):
        return RunStatus.WARN
    for name in ("L2", "L3"):
        result = bundle.levels.get(name)
        if result is not None and not result.passed:
            return RunStatus.WARN
    if bundle.render is not None and not getattr(bundle.render, "within_tolerance", True):
        return RunStatus.WARN
    return RunStatus.PASS
