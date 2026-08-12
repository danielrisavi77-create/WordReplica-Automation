from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

import pytest

from scripts.persistent_harness.layout import PersistentLayout
from scripts.persistent_harness.updater import UpdateTransactionError, apply_update, recover_incomplete_transaction
from scripts.persistent_harness.versioning import VersionMetadata


def tree_hash(root: Path) -> str:
    h = sha256()
    if not root.exists():
        return h.hexdigest()
    for path in sorted(p for p in root.rglob("*") if p.is_file()):
        h.update(path.relative_to(root).as_posix().encode())
        h.update(path.read_bytes())
    return h.hexdigest()


def seed_layout(tmp_path: Path) -> PersistentLayout:
    layout = PersistentLayout.from_root(tmp_path)
    for path in (layout.input_dir, layout.results_root, layout.current, layout.backup, layout.updates):
        path.mkdir(parents=True, exist_ok=True)
    (layout.input_dir / "a.docx").write_bytes(b"source")
    (layout.results_root / "old" / "summary.json").parent.mkdir(parents=True)
    (layout.results_root / "old" / "summary.json").write_text("old", encoding="utf-8")
    (layout.current / "app.txt").write_text("old-code", encoding="utf-8")
    (layout.current / "pyproject.toml").write_text("[project]\nname='x'\nversion='1'\ndependencies=[]\n", encoding="utf-8")
    layout.config.write_text("{}", encoding="utf-8")
    VersionMetadata(1, "2.0.0", "2026-08-11T19:00:00+02:00", "a" * 64, None).write_atomic(layout.version_file)
    return layout


def update_zip(tmp_path: Path, content="new-code", target="2.1.0") -> Path:
    payload = {"app.txt": content.encode(), "pyproject.toml": b"[project]\nname='x'\nversion='1'\ndependencies=[]\n"}
    manifest = {
        "schema_version": 1,
        "package_id": f"update-{target}",
        "target_version": target,
        "minimum_installed_version": "2.0.0",
        "files": {name: sha256(data).hexdigest() for name, data in payload.items()},
    }
    path = tmp_path / f"update-{target}.zip"
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("update_manifest.json", json.dumps(manifest))
        for name, data in payload.items():
            zf.writestr(f"payload/{name}", data)
    return path


def test_successful_update_never_changes_input_or_results(tmp_path):
    layout = seed_layout(tmp_path)
    input_before, results_before = tree_hash(layout.input_dir), tree_hash(layout.results_root)
    meta = apply_update(layout, update_zip(tmp_path), smoke_test=lambda current: None)
    assert (layout.current / "app.txt").read_text() == "new-code"
    assert meta.harness_version == "2.1.0"
    assert tree_hash(layout.input_dir) == input_before
    assert tree_hash(layout.results_root) == results_before
    assert any(layout.backup.iterdir())
    assert not (layout.root / ".update_transaction.json").exists()


def test_failed_preflight_leaves_current_untouched(tmp_path):
    layout = seed_layout(tmp_path)
    current_before = tree_hash(layout.current)
    with pytest.raises(UpdateTransactionError, match="pre-install validation"):
        apply_update(layout, update_zip(tmp_path), smoke_test=lambda current: (_ for _ in ()).throw(RuntimeError("bad staging")))
    assert tree_hash(layout.current) == current_before
    assert VersionMetadata.load(layout.version_file).harness_version == "2.0.0"


def test_post_install_failure_restores_previous_current(tmp_path):
    layout = seed_layout(tmp_path)
    current_before = tree_hash(layout.current)
    calls = []
    def smoke(current):
        current = Path(current)
        calls.append(current)
        if (current / "app.txt").read_text(encoding="utf-8") == "new-code" and current.resolve() == layout.current.resolve():
            raise RuntimeError("bad live")
    with pytest.raises(UpdateTransactionError, match="post-install validation"):
        apply_update(layout, update_zip(tmp_path), smoke_test=smoke)
    assert tree_hash(layout.current) == current_before
    assert VersionMetadata.load(layout.version_file).harness_version == "2.0.0"
    assert not (layout.root / ".update_transaction.json").exists()


def test_recovery_restores_known_good_backup_after_interrupted_swap(tmp_path):
    layout = seed_layout(tmp_path)
    backup_dir = layout.backup / "20260811-190000-2.0.0"
    (backup_dir / "current").mkdir(parents=True)
    (backup_dir / "current" / "app.txt").write_text("known-good", encoding="utf-8")
    VersionMetadata(1, "2.0.0", "2026-08-11T19:00:00+02:00", "a" * 64, None).write_atomic(backup_dir / "version.json")
    (layout.current / "app.txt").write_text("half-installed", encoding="utf-8")
    marker = {"schema_version": 1, "phase": "swapped", "backup_dir": str(backup_dir)}
    (layout.root / ".update_transaction.json").write_text(json.dumps(marker), encoding="utf-8")
    recovered = recover_incomplete_transaction(layout, smoke_test=lambda current: None)
    assert recovered is True
    assert (layout.current / "app.txt").read_text() == "known-good"
    assert VersionMetadata.load(layout.version_file).harness_version == "2.0.0"
    assert not (layout.root / ".update_transaction.json").exists()


def test_two_updates_preserve_same_input_and_old_results(tmp_path):
    layout = seed_layout(tmp_path)
    input_before = tree_hash(layout.input_dir)
    results_before = tree_hash(layout.results_root)
    apply_update(layout, update_zip(tmp_path, content="v21", target="2.1.0"), smoke_test=lambda current: None)
    apply_update(layout, update_zip(tmp_path, content="v22", target="2.2.0"), smoke_test=lambda current: None)
    assert (layout.current / "app.txt").read_text() == "v22"
    assert tree_hash(layout.input_dir) == input_before
    assert tree_hash(layout.results_root) == results_before
    assert len([p for p in layout.backup.iterdir() if p.is_dir()]) == 2


def test_manual_rollback_does_not_touch_input_or_results(tmp_path):
    from scripts.persistent_harness.rollback import rollback_latest
    layout = seed_layout(tmp_path)
    input_before = tree_hash(layout.input_dir)
    results_before = tree_hash(layout.results_root)
    apply_update(layout, update_zip(tmp_path), smoke_test=lambda current: None)
    restored = rollback_latest(layout, smoke_test=lambda current: None)
    assert restored.harness_version == "2.0.0"
    assert (layout.current / "app.txt").read_text() == "old-code"
    assert tree_hash(layout.input_dir) == input_before
    assert tree_hash(layout.results_root) == results_before
