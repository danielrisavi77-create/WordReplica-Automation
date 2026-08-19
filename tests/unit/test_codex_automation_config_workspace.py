from pathlib import Path

from scripts.codex_automation.config import CodexAutomationConfig
from scripts.codex_automation.workspace import GoldenWorkspace, sha256_file


def test_default_golden_timeout_covers_long_complex_documents():
    assert CodexAutomationConfig().reconstruction_timeout_seconds == 14_400

    from scripts.codex_automation.config import load_config

    repo_config = Path(__file__).resolve().parents[2] / "codex_automation.json"
    assert load_config(repo_config).reconstruction_timeout_seconds == 14_400


def test_workspace_copies_golden_without_mutating_source(tmp_path):
    root = tmp_path / "root"
    golden_dir = root / "golden"
    golden_dir.mkdir(parents=True)
    golden = golden_dir / "golden.docx"
    golden.write_bytes(b"golden-bytes")
    config = CodexAutomationConfig(local_root=root, golden_filename="golden.docx")

    workspace = GoldenWorkspace(config)
    run = workspace.create_run("abc12345")

    assert run.source_copy.read_bytes() == b"golden-bytes"
    assert run.source_copy != golden
    assert run.source_sha256_before == sha256_file(golden)
    assert workspace.verify_golden_unchanged(run) is True


def test_workspace_lock_allows_only_one_active_run(tmp_path):
    root = tmp_path / "root"
    (root / "golden").mkdir(parents=True)
    (root / "golden" / "golden.docx").write_bytes(b"x")
    config = CodexAutomationConfig(local_root=root, golden_filename="golden.docx")
    first = GoldenWorkspace(config)
    second = GoldenWorkspace(config)

    first.acquire_lock()
    try:
        try:
            second.acquire_lock()
        except RuntimeError as exc:
            assert "already active" in str(exc)
        else:
            raise AssertionError("second lock unexpectedly succeeded")
    finally:
        first.release_lock()


def test_golden_documents_lists_primary_first_then_additional_by_position():
    config = CodexAutomationConfig(
        golden_filename="a.docx", additional_golden_filenames=["b.docx", "c.docx"],
    )

    assert config.golden_documents == [("golden_1", "a.docx"), ("golden_2", "b.docx"), ("golden_3", "c.docx")]
    assert config.golden_path_for("golden_2") == config.golden_dir / "b.docx"
    assert config.golden_path_for("golden_1") == config.golden_path


def test_golden_path_for_unknown_id_raises():
    config = CodexAutomationConfig()
    try:
        config.golden_path_for("golden_9")
    except KeyError as exc:
        assert "golden_9" in str(exc)
    else:
        raise AssertionError("expected KeyError for unknown golden id")


def test_workspace_create_run_targets_requested_golden_document(tmp_path):
    root = tmp_path / "root"
    golden_dir = root / "golden"
    golden_dir.mkdir(parents=True)
    (golden_dir / "golden.docx").write_bytes(b"primary")
    (golden_dir / "second.docx").write_bytes(b"secondary")
    config = CodexAutomationConfig(
        local_root=root, golden_filename="golden.docx", additional_golden_filenames=["second.docx"],
    )

    workspace = GoldenWorkspace(config)
    run = workspace.create_run("abc12345", golden_id="golden_2")

    assert run.golden_id == "golden_2"
    assert run.source_copy.read_bytes() == b"secondary"
    assert workspace.verify_golden_unchanged(run) is True


def test_load_config_reads_repo_json_and_optional_root_override(tmp_path):
    import json
    from scripts.codex_automation.config import load_config

    cfg = tmp_path / "codex_automation.json"
    cfg.write_text(json.dumps({
        "local_root": "C:/Ignored",
        "golden_filename": "x.docx",
        "reconstruction_timeout_seconds": 123,
    }), encoding="utf-8")

    loaded = load_config(cfg, local_root_override=tmp_path / "actual")

    assert loaded.local_root == tmp_path / "actual"
    assert loaded.golden_filename == "x.docx"
    assert loaded.reconstruction_timeout_seconds == 123


def _locked_workspace(tmp_path):
    root = tmp_path / "root"
    (root / "golden").mkdir(parents=True)
    (root / "golden" / "golden.docx").write_bytes(b"x")
    config = CodexAutomationConfig(local_root=root, golden_filename="golden.docx")
    return GoldenWorkspace(config), root / "state" / "golden_run.lock"


def _write_lock(path, record):
    import json

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(record), encoding="utf-8")


def test_lock_records_the_holder_process_identity(tmp_path):
    # Staleness can only be decided if the lock says who holds it, by both pid
    # and process creation FILETIME -- pid alone is reusable.
    import json
    import os

    workspace, lock_path = _locked_workspace(tmp_path)
    workspace.acquire_lock()
    try:
        record = json.loads(lock_path.read_text(encoding="utf-8"))
    finally:
        workspace.release_lock()

    assert record["pid"] == os.getpid()
    assert isinstance(record["creation_filetime"], int)


def test_lock_left_by_a_dead_process_is_reclaimed(tmp_path):
    # A crashed run must not block every future Golden run forever.
    workspace, lock_path = _locked_workspace(tmp_path)
    _write_lock(lock_path, {"pid": 424242, "creation_filetime": 999, "created_utc": "x"})

    workspace.acquire_lock(identity_probe=lambda pid: None)
    try:
        assert lock_path.is_file()
    finally:
        workspace.release_lock()


def test_lock_held_by_a_live_process_is_never_reclaimed(tmp_path):
    workspace, lock_path = _locked_workspace(tmp_path)
    _write_lock(lock_path, {"pid": 4242, "creation_filetime": 777, "created_utc": "x"})

    try:
        workspace.acquire_lock(identity_probe=lambda pid: 777)
    except RuntimeError as exc:
        assert "4242" in str(exc)
    else:
        workspace.release_lock()
        raise AssertionError("lock held by a live process was reclaimed")


def test_reused_pid_does_not_count_as_a_live_holder(tmp_path):
    # The pid is alive but it is a different process than the one that took the
    # lock -- exactly the case pid-only checks get wrong.
    workspace, lock_path = _locked_workspace(tmp_path)
    _write_lock(lock_path, {"pid": 4242, "creation_filetime": 777, "created_utc": "x"})

    workspace.acquire_lock(identity_probe=lambda pid: 111)
    try:
        assert lock_path.is_file()
    finally:
        workspace.release_lock()


def test_unprovable_holder_identity_keeps_the_lock(tmp_path):
    # If the platform cannot answer, the safe answer is "held". Never clear a
    # lock we cannot prove is stale.
    workspace, lock_path = _locked_workspace(tmp_path)
    _write_lock(lock_path, {"pid": 4242, "creation_filetime": 777, "created_utc": "x"})

    def _cannot_tell(pid):
        raise RuntimeError("Windows process identity is unavailable")

    try:
        workspace.acquire_lock(identity_probe=_cannot_tell)
    except RuntimeError as exc:
        assert "already active" in str(exc)
    else:
        workspace.release_lock()
        raise AssertionError("lock was reclaimed without proof it was stale")


def test_unreadable_lock_record_keeps_the_lock(tmp_path):
    workspace, lock_path = _locked_workspace(tmp_path)
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    lock_path.write_text("not json", encoding="utf-8")

    try:
        workspace.acquire_lock(identity_probe=lambda pid: None)
    except RuntimeError as exc:
        assert "already active" in str(exc)
    else:
        workspace.release_lock()
        raise AssertionError("lock with an unreadable record was reclaimed")


def test_inaccessible_process_is_unknown_not_dead(monkeypatch):
    # OpenProcess fails with ERROR_ACCESS_DENIED for a process that very much
    # exists. Reading that as "gone" would clear a live holder's lock -- the one
    # direction this check must never get wrong.
    from scripts.codex_automation import workspace as ws

    def _access_denied(pid):
        raise OSError(5, "Access is denied")

    monkeypatch.setattr(ws, "process_creation_filetime", _access_denied, raising=False)
    import pytest

    with pytest.raises(RuntimeError):
        ws._process_identity(4242)


def test_missing_process_is_reported_as_gone(monkeypatch):
    from scripts.codex_automation import workspace as ws

    def _no_such_process(pid):
        raise OSError(87, "The parameter is incorrect")

    monkeypatch.setattr(ws, "process_creation_filetime", _no_such_process, raising=False)
    assert ws._process_identity(4242) is None
