"""Post-verification handoff of a repair-poc output to the user.

Only call open_output_for_user after a RepairCompletionReport shows
full_pass=True. This never obtains, activates, or closes any Word instance
itself — it just asks Windows to open the file with its default
association, exactly like a user double-clicking it in Explorer, so it
cannot interfere with an unrelated already-open Word document.
"""
from __future__ import annotations

import os
from pathlib import Path


def open_output_for_user(path: Path) -> None:
    os.startfile(str(path))
