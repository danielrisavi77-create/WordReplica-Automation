from __future__ import annotations

import argparse
from pathlib import Path
import sys

CURRENT_ROOT = Path(__file__).resolve().parents[2]
for item in (str(CURRENT_ROOT), str(CURRENT_ROOT / "src")):
    if item not in sys.path:
        sys.path.insert(0, item)

from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.rollback import rollback_latest
from scripts.persistent_harness.smoke import compile_smoke


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    args = parser.parse_args(argv)
    layout = PersistentLayout.from_root(Path(args.root))
    try:
        meta = rollback_latest(layout, compile_smoke)
    except Exception as exc:
        print(f"HARNESS ROLLBACK: FAIL: {exc}")
        return 2
    print(f"HARNESS ROLLBACK: PASS: restored {meta.harness_version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
