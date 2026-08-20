"""Asking Microsoft Word whether it accepts a reconstruction at all.

Every gate in this project compares a reconstruction against its source. This
asks a different and blunter question: does Word open the document, or does it
show the "unreadable content" prompt and repair it? A document Word repairs has
failed regardless of how well its projections match.

Two rules make the answer usable:

* **Only a repair the source did not have is ours.** Corpus documents include
  packages Word already repairs -- deliberately malformed regression fixtures,
  clusterfuzz output. Counting those against the reconstruction would drown the
  signal in defects that were there before we touched anything.
* **"Not verified" is not "clean".** ``detect_open_and_repair`` returns None
  when the probe could not complete -- no Word, a COM timeout, a hung process.
  Reading that as a pass is how a fidelity claim quietly becomes untrue.

The lab always yields the machine-wide Word lock to the Golden pipeline. It
never clears a lock it finds: deciding a lock is stale needs the holder's
process identity, which is ``GoldenWorkspace``'s job, not this module's.
"""
from __future__ import annotations

from enum import StrEnum
from pathlib import Path

__all__ = ["RepairOutcome", "classify_repair", "lock_is_held", "probe_repair"]

DEFAULT_LOCK_PATH = Path("C:/WordReplica-Automation/state/golden_run.lock")


class RepairOutcome(StrEnum):
    CLEAN = "CLEAN"                          # Word opened both without repairing
    REPAIR_INTRODUCED = "REPAIR_INTRODUCED"  # ours: the source was fine
    REPAIR_INHERITED = "REPAIR_INHERITED"    # the source was already repaired
    REPAIR_RESOLVED = "REPAIR_RESOLVED"      # the rebuild fixed a broken source
    UNVERIFIED = "UNVERIFIED"                # the probe could not complete


def classify_repair(source: bool | None, output: bool | None) -> RepairOutcome:
    """Turn two probe answers into one verdict about the reconstruction."""
    if source is None or output is None:
        return RepairOutcome.UNVERIFIED
    if not source and not output:
        return RepairOutcome.CLEAN
    if not source and output:
        return RepairOutcome.REPAIR_INTRODUCED
    if source and output:
        return RepairOutcome.REPAIR_INHERITED
    return RepairOutcome.REPAIR_RESOLVED


def lock_is_held(lock_path: Path = DEFAULT_LOCK_PATH) -> bool:
    """Whether a Golden Word run holds the machine-wide lock.

    Looks only. A lock left behind by a crashed run is reclaimed by
    ``GoldenWorkspace``, which can prove the holder is gone from its recorded
    pid and process creation time; guessing here would risk two Word runs at
    once, which ``AGENTS.md`` forbids.
    """
    return Path(lock_path).exists()


def probe_repair(source: Path, output: Path, *, detector=None) -> RepairOutcome:
    """Ask Word about the reconstruction, and about the source only if needed.

    An open costs roughly 25 seconds on this machine, so probing both documents
    unconditionally doubles the cost of the whole lane. When Word opens the
    reconstruction cleanly there is nothing the source could add -- the question
    being asked is whether we hand back something Word repairs. The source is
    opened only when the output was repaired, to find out whose defect it is.
    """
    if detector is None:
        from word_replica.qa.word_render import detect_open_and_repair as detector

    try:
        output_repaired = detector(Path(output))
        if output_repaired is False:
            return RepairOutcome.CLEAN
        source_repaired = detector(Path(source))
    except Exception:
        # A probe that raised is a probe that did not answer. Never a pass.
        return RepairOutcome.UNVERIFIED
    return classify_repair(source_repaired, output_repaired)
