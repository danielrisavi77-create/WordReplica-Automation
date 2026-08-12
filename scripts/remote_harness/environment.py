from __future__ import annotations

from datetime import datetime, timezone
import locale
import os
from pathlib import Path
import platform
import shutil
import sys

from scripts.remote_harness import HARNESS_VERSION, SCHEMA_VERSION
from scripts.remote_harness.prerequisites import _default_word_probe


def capture_environment(root: Path) -> dict:
    word_ok, word_info = _default_word_probe() if os.name == "nt" else (False, {"error":"not Windows"})
    usage=shutil.disk_usage(Path(root))
    return {
        "schema_version": SCHEMA_VERSION,
        "harness_version": HARNESS_VERSION,
        "captured_at_utc": datetime.now(timezone.utc).isoformat(),
        "windows": {
            "platform": platform.platform(),
            "release": platform.release(),
            "version": platform.version(),
            "machine": platform.machine(),
        },
        "python": {"version": sys.version, "executable": sys.executable, "architecture": platform.architecture()[0]},
        "locale": locale.getlocale(),
        "word": {"available": word_ok, **word_info},
        "disk": {"free_mb": int(usage.free//(1024*1024)), "total_mb": int(usage.total//(1024*1024))},
        "flags": {"WORD_REPLICA_WORD_TESTS": os.environ.get("WORD_REPLICA_WORD_TESTS")},
    }
