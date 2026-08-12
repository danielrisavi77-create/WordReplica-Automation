from pathlib import Path

from scripts.persistent_harness.dependencies import dependency_fingerprint, dependencies_changed


def write_pyproject(root: Path, dependency: str) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    path = root / "pyproject.toml"
    path.write_text(
        "[project]\nname='x'\nversion='1'\nrequires-python='>=3.12'\ndependencies=[\"" + dependency + "\"]\n",
        encoding="utf-8",
    )
    return path


def test_unchanged_dependency_fingerprint_reuses_venv(tmp_path):
    old = dependency_fingerprint(write_pyproject(tmp_path / "old", "lxml>=5"))
    new = dependency_fingerprint(write_pyproject(tmp_path / "new", "lxml>=5"))
    assert old == new
    assert dependencies_changed(old, new) is False


def test_changed_dependency_fingerprint_requires_sync(tmp_path):
    old = dependency_fingerprint(write_pyproject(tmp_path / "old", "lxml>=5"))
    new = dependency_fingerprint(write_pyproject(tmp_path / "new", "lxml>=6"))
    assert old != new
    assert dependencies_changed(old, new) is True
