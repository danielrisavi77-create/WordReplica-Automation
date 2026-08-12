from word_replica.domain.enums import RunStatus
from word_replica.domain.results import WarningItem
from word_replica.qa.policy import QaBundle, QaLevelResult, classify_run


def test_content_loss_is_fail():
    bundle = QaBundle(levels={"L0": QaLevelResult("L0", passed=False, findings=["text mismatch"])})
    assert classify_run(bundle, warnings=[]) is RunStatus.FAIL


def test_noncritical_visual_warning_is_warn():
    bundle = QaBundle(
        levels={
            "L0": QaLevelResult("L0", True, []),
            "L1": QaLevelResult("L1", True, []),
            "L2": QaLevelResult("L2", True, []),
            "L3": QaLevelResult("L3", True, []),
        }
    )
    assert classify_run(bundle, warnings=[WarningItem("VISUAL_DIFF", "minor mismatch")]) is RunStatus.WARN


def test_required_structure_loss_is_fail():
    bundle = QaBundle(
        levels={
            "L0": QaLevelResult("L0", True, []),
            "L1": QaLevelResult("L1", False, ["asset loss"], has_required_loss=True),
        }
    )
    assert classify_run(bundle, warnings=[]) is RunStatus.FAIL
