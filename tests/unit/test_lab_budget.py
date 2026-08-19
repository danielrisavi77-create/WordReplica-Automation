"""The disk governor.

This machine has had under five gigabytes free while a Golden run was in
flight. An unbounded lab fetch would not just fail the lab; it would break
RUN_GOLDEN_CODEX.ps1, which is the project's real gate.
"""
from dataclasses import dataclass

import pytest

from word_replica.lab.budget import DiskBudget, DiskExhausted


@dataclass
class _Usage:
    free: int


def _budget(free_bytes: int, **kwargs) -> DiskBudget:
    return DiskBudget(usage=lambda _path: _Usage(free=free_bytes), **kwargs)


def test_a_batch_refuses_to_start_below_the_floor(tmp_path):
    with pytest.raises(DiskExhausted, match="refusing to start"):
        _budget(1 * 1024 ** 3).assert_can_start(tmp_path)


def test_a_batch_starts_when_there_is_room(tmp_path):
    _budget(5 * 1024 ** 3).assert_can_start(tmp_path)


def test_a_running_batch_aborts_at_the_lower_threshold(tmp_path):
    # The gap between the two thresholds is what lets a batch that legitimately
    # started still finish and flush.
    budget = _budget(1 * 1024 ** 3)
    budget.assert_can_continue(tmp_path)

    with pytest.raises(DiskExhausted, match="aborting mid-batch"):
        _budget(100 * 1024 ** 2).assert_can_continue(tmp_path)


def test_headroom_accounts_for_what_is_about_to_be_written(tmp_path):
    budget = _budget(1 * 1024 ** 3)

    assert budget.headroom_for(tmp_path, 100 * 1024 ** 2) is True
    assert budget.headroom_for(tmp_path, 900 * 1024 ** 2) is False


def test_free_space_is_probed_on_an_ancestor_when_the_path_does_not_exist_yet(tmp_path):
    probed: list[str] = []

    def _usage(path: str):
        probed.append(path)
        return _Usage(free=5 * 1024 ** 3)

    DiskBudget(usage=_usage).assert_can_start(tmp_path / "not" / "created" / "yet")

    assert probed == [str(tmp_path)]


def test_the_governor_reads_real_free_space(tmp_path):
    # No mock: the default must actually work, or every guard above is theatre.
    assert DiskBudget().free_bytes(tmp_path) > 0
