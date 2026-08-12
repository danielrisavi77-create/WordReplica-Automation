from __future__ import annotations

from pathlib import Path
import subprocess
import sys


def compile_smoke(current: Path) -> None:
    current = Path(current).resolve()
    required = [
        current / "pyproject.toml",
        current / "RUN_REMOTE_WORD_HARNESS.ps1",
        current / "scripts" / "remote_harness" / "main.py",
        current / "scripts" / "persistent_harness" / "update_cli.py",
        current / "scripts" / "persistent_harness" / "rollback_cli.py",
        current / "src" / "word_replica" / "__init__.py",
    ]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise RuntimeError("required update payload files missing: " + ", ".join(missing))
    proc = subprocess.run(
        [sys.executable, "-m", "compileall", "-q", str(current / "src"), str(current / "scripts")],
        cwd=current,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError((proc.stderr or proc.stdout or "compileall failed").strip())


def runtime_smoke(current: Path) -> None:
    current = Path(current).resolve()
    code = (
        "import sys; "
        f"sys.path.insert(0, {str(current)!r}); "
        f"sys.path.insert(0, {str(current / 'src')!r}); "
        "import word_replica; import scripts.remote_harness.main; print('OK')"
    )
    proc = subprocess.run([sys.executable, "-c", code], cwd=current, capture_output=True, text=True)
    if proc.returncode != 0 or "OK" not in proc.stdout:
        raise RuntimeError((proc.stderr or proc.stdout or "runtime import smoke failed").strip())
