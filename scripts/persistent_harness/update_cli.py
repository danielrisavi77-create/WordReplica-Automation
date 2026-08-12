from __future__ import annotations

import argparse
from pathlib import Path
import subprocess
import sys

CURRENT_ROOT = Path(__file__).resolve().parents[2]
for item in (str(CURRENT_ROOT), str(CURRENT_ROOT / "src")):
    if item not in sys.path:
        sys.path.insert(0, item)

from scripts.persistent_harness.dependencies import dependency_fingerprint, dependencies_changed
from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.rollback import rollback_latest
from scripts.persistent_harness.smoke import compile_smoke, runtime_smoke
from scripts.persistent_harness.updater import apply_update


def _pip_sync(python: Path, current: Path) -> None:
    proc = subprocess.run([str(python), "-m", "pip", "install", "-e", str(current), "--no-cache-dir"])
    if proc.returncode != 0:
        raise RuntimeError(f"dependency synchronization failed with exit code {proc.returncode}")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--zip", required=True)
    parser.add_argument("--venv-python", required=True)
    args = parser.parse_args(argv)
    layout = PersistentLayout.from_root(Path(args.root))
    python = Path(args.venv_python)
    old_fp = dependency_fingerprint(layout.current / "pyproject.toml")
    try:
        meta = apply_update(layout, Path(args.zip), smoke_test=compile_smoke)
        new_fp = dependency_fingerprint(layout.current / "pyproject.toml")
        if dependencies_changed(old_fp, new_fp):
            try:
                _pip_sync(python, layout.current)
                runtime_smoke(layout.current)
            except Exception as exc:
                rollback_latest(layout, compile_smoke)
                _pip_sync(python, layout.current)
                raise RuntimeError(f"update rolled back after dependency/runtime failure: {exc}") from exc
        else:
            runtime_smoke(layout.current)
    except Exception as exc:
        print(f"HARNESS UPDATE: FAIL: {exc}")
        return 2
    print(f"HARNESS UPDATE: PASS: {meta.previous_version} -> {meta.harness_version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
