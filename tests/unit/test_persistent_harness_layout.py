from pathlib import Path

from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.versioning import VersionMetadata


def test_layout_keeps_user_data_outside_current(tmp_path: Path):
    layout = PersistentLayout.from_root(tmp_path)
    assert layout.input_dir == tmp_path.resolve() / "realworld_input"
    assert layout.results_root == tmp_path.resolve() / "results"
    assert layout.current == tmp_path.resolve() / "current"
    assert layout.venv == tmp_path.resolve() / ".venv_harness"
    assert layout.config == tmp_path.resolve() / "harness_config.json"


def test_version_metadata_round_trips_atomically(tmp_path: Path):
    path = tmp_path / "version.json"
    value = VersionMetadata(1, "2.0.0", "2026-08-11T19:00:00+02:00", "a" * 64, "1.0.0")
    value.write_atomic(path)
    assert VersionMetadata.load(path) == value
    assert not (tmp_path / "version.json.tmp").exists()
