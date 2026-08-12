from pathlib import Path
from word_replica.config import RebuildOptions
from word_replica.services.project_store import ProjectStore


def test_create_project_builds_required_directories(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"docx")
    store = ProjectStore(app_root=tmp_path / "app")
    paths = store.create_project(source, RebuildOptions())
    assert paths.source_snapshot_dir.is_dir()
    assert paths.working_dir.is_dir()
    assert paths.output_dir.is_dir()
    assert paths.backups_dir.is_dir()
    assert paths.logs_dir.is_dir()
    assert paths.qa_dir.is_dir()
    assert paths.project_json.exists()
    assert store.get_project(paths.project_id)["source_path"] == str(source.resolve())


def test_paths_for_project_rehydrates_existing_project_directories(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"docx")
    store = ProjectStore(app_root=tmp_path / "app")
    created = store.create_project(source, RebuildOptions())
    restored = store.paths_for_project(created.project_id)
    assert restored == created


def test_list_resumable_projects_only_returns_checkpointed_interactive_stopped_projects(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"docx")
    store = ProjectStore(app_root=tmp_path / "app")
    stopped = store.create_project(source, RebuildOptions())
    completed = store.create_project(source, RebuildOptions())
    store.set_status(stopped.project_id, "STOPPED")
    store.set_status(completed.project_id, "COMPLETED")
    checkpoint = stopped.logs_dir / "interactive_checkpoint.json"
    checkpoint.write_text('{"status":"STOPPED","last_completed_event_index":7,"timestamp_utc":"2026-08-11T12:00:00+00:00"}', encoding="utf-8")

    items = store.list_resumable_projects()

    assert [item["id"] for item in items] == [stopped.project_id]
    assert items[0]["checkpoint_path"] == str(checkpoint)
    assert items[0]["last_completed_event_index"] == 7
