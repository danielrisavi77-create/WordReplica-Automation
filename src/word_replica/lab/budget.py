"""Disk governor.

The lab shares a single writable drive with the repository, the Golden
pipeline's diagnostics, and Microsoft Word's own temporary files. Free space
has been observed at under five gigabytes while a Golden run was in flight, so
an unbounded fetch would not merely fail the lab -- it would break
``RUN_GOLDEN_CODEX.ps1``, which is the project's actual gate.

Two rules follow, and both are enforced rather than documented:

* Check headroom **before** every write, not after. A guard that notices
  afterwards has already caused the problem.
* Never delete anything the lab does not own. If the floor cannot be met, stop
  and say so; reclaiming space from ``diagnostics/`` or ``golden/`` is not the
  lab's call to make.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
from typing import Callable

__all__ = ["DiskBudget", "DiskExhausted"]


class DiskExhausted(RuntimeError):
    """Raised instead of writing when free space is below the floor."""


@dataclass(frozen=True, slots=True)
class DiskBudget:
    """Free-space floors for the drive the lab writes to.

    ``floor_bytes`` is where a batch refuses to start; ``abort_bytes`` is where
    one already running gives up mid-way. The gap exists so a batch that starts
    legitimately still has room to finish and flush.
    """

    floor_bytes: int = 2 * 1024 ** 3
    abort_bytes: int = 800 * 1024 ** 2
    usage: Callable[[str], object] = shutil.disk_usage

    def free_bytes(self, path: Path) -> int:
        probe = Path(path)
        while not probe.exists() and probe != probe.parent:
            probe = probe.parent
        return int(self.usage(str(probe)).free)

    def assert_can_start(self, path: Path) -> None:
        free = self.free_bytes(path)
        if free < self.floor_bytes:
            raise DiskExhausted(
                f"refusing to start: {free / 1024**3:.2f} GB free at {path}, "
                f"floor is {self.floor_bytes / 1024**3:.2f} GB"
            )

    def assert_can_continue(self, path: Path) -> None:
        free = self.free_bytes(path)
        if free < self.abort_bytes:
            raise DiskExhausted(
                f"aborting mid-batch: {free / 1024**3:.2f} GB free at {path}, "
                f"abort threshold is {self.abort_bytes / 1024**3:.2f} GB"
            )

    def headroom_for(self, path: Path, wanted_bytes: int) -> bool:
        """True when writing ``wanted_bytes`` would still leave the floor intact."""
        return self.free_bytes(path) - wanted_bytes >= self.abort_bytes
