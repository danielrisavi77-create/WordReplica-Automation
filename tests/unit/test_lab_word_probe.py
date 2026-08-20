"""The strongest single property the lab can assert: Word does not repair us.

Every gate so far compares a reconstruction against its source. This asks a
different question -- does Microsoft Word accept the document at all, or does it
show the "unreadable content" prompt? A reconstruction Word repairs has failed
regardless of how well its projections match, and none of the package-level
work in this project had been checked against real Word until now.

The probe itself already exists (`qa/word_render.detect_open_and_repair`) and
returns None when it could not complete. None must never be read as a pass:
"not verified" and "verified clean" are different answers, and conflating them
is exactly how a fidelity claim becomes untrue.
"""
from pathlib import Path

import pytest

from word_replica.lab.word_probe import (
    RepairOutcome,
    classify_repair,
    lock_is_held,
)


# --- what the probe's three answers mean --------------------------------------

@pytest.mark.parametrize(
    "source, output, expected",
    [
        (False, False, RepairOutcome.CLEAN),
        (False, True, RepairOutcome.REPAIR_INTRODUCED),
        (True, True, RepairOutcome.REPAIR_INHERITED),
        (True, False, RepairOutcome.REPAIR_RESOLVED),
    ],
)
def test_a_repair_is_only_ours_if_the_source_did_not_have_one(source, output, expected):
    # A source Word already repairs is not evidence against the reconstruction.
    # Separating the two is what makes the number actionable.
    assert classify_repair(source, output) is expected


@pytest.mark.parametrize("source, output", [(None, False), (False, None), (None, None)])
def test_an_incomplete_probe_is_unverified_not_clean(source, output):
    assert classify_repair(source, output) is RepairOutcome.UNVERIFIED


def test_unverified_does_not_count_as_a_repair():
    # Fail-closed on the claim, not on the document: an unfinished probe must
    # neither pass the reconstruction nor accuse it.
    outcome = classify_repair(None, None)

    assert outcome is not RepairOutcome.CLEAN
    assert outcome is not RepairOutcome.REPAIR_INTRODUCED


# --- yielding to the Golden pipeline ------------------------------------------

def test_the_lab_yields_when_a_golden_run_holds_the_lock(tmp_path):
    # AGENTS.md: one active Golden Word run at a time, machine-wide. The lab
    # always gives way -- RUN_GOLDEN_CODEX.ps1 must never wait on it.
    lock = tmp_path / "golden_run.lock"
    lock.write_text('{"pid": 4242}', encoding="utf-8")

    assert lock_is_held(lock) is True


def test_no_lock_means_the_lab_may_use_word(tmp_path):
    assert lock_is_held(tmp_path / "golden_run.lock") is False


def test_the_lab_never_clears_the_lock_itself(tmp_path):
    # Reclaiming a stale lock is GoldenWorkspace's decision, made with the
    # holder's process identity. The lab only ever looks.
    lock = tmp_path / "golden_run.lock"
    lock.write_text('{"pid": 4242}', encoding="utf-8")

    lock_is_held(lock)

    assert lock.exists()


# --- the source is only opened when the answer depends on it ------------------

def test_a_clean_output_does_not_cost_a_second_word_open(tmp_path):
    # Measured at roughly 25 seconds per open on this machine, so probing both
    # documents unconditionally doubles the cost of the whole lane. When Word
    # opens the reconstruction cleanly there is nothing the source could add:
    # the question the lane asks is whether we hand back something Word repairs.
    from word_replica.lab import word_probe

    opened = []

    def _probe(path):
        opened.append(Path(path).name)
        return False

    outcome = word_probe.probe_repair(
        tmp_path / "source.docx", tmp_path / "output.docx", detector=_probe
    )

    assert outcome is RepairOutcome.CLEAN
    assert opened == ["output.docx"]


def test_a_repaired_output_costs_a_second_open_to_attribute_blame(tmp_path):
    # Corpus documents include packages Word already repairs. Counting those
    # against the reconstruction would drown the signal, so when the output is
    # repaired the source is opened to find out whose defect it is.
    from word_replica.lab import word_probe

    opened = []

    def _probe(path):
        opened.append(Path(path).name)
        return True  # both repaired

    outcome = word_probe.probe_repair(
        tmp_path / "source.docx", tmp_path / "output.docx", detector=_probe
    )

    assert outcome is RepairOutcome.REPAIR_INHERITED
    assert opened == ["output.docx", "source.docx"]


def test_a_repair_we_introduced_is_still_identified(tmp_path):
    from word_replica.lab import word_probe

    def _probe(path):
        return Path(path).name == "output.docx"

    outcome = word_probe.probe_repair(
        tmp_path / "source.docx", tmp_path / "output.docx", detector=_probe
    )

    assert outcome is RepairOutcome.REPAIR_INTRODUCED


def test_a_probe_that_raises_is_unverified(tmp_path):
    from word_replica.lab import word_probe

    def _explode(path):
        raise RuntimeError("COM went away")

    outcome = word_probe.probe_repair(
        tmp_path / "source.docx", tmp_path / "output.docx", detector=_explode
    )

    assert outcome is RepairOutcome.UNVERIFIED
