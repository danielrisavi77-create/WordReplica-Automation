from pathlib import Path


def test_persistent_launcher_passes_protected_paths_and_version():
    text = Path("persistent_root/RUN_PERSISTENT_HARNESS.ps1").read_text(encoding="ascii")
    assert "realworld_input" in text
    assert "results" in text
    assert "version.json" in text
    assert "-InputDir" in text and "-ResultsRoot" in text and "-HarnessVersion" in text
    assert ".venv_harness" in text


def test_persistent_scripts_are_ascii_safe():
    for name in ("START_HERE.cmd", "RUN_PERSISTENT_HARNESS.ps1", "IMPORT_EXISTING_INPUT.cmd", "IMPORT_EXISTING_INPUT.ps1"):
        raw = (Path("persistent_root") / name).read_bytes()
        assert raw
        assert all(byte < 128 for byte in raw), name


def test_update_and_rollback_scripts_are_ascii_safe_and_protect_user_data():
    for name in ("UPDATE_HARNESS.cmd", "UPDATE_HARNESS.ps1", "ROLLBACK_HARNESS.cmd", "ROLLBACK_HARNESS.ps1"):
        path = Path("persistent_root") / name
        raw = path.read_bytes()
        assert raw
        assert all(byte < 128 for byte in raw), name
        text = raw.decode("ascii")
        assert "realworld_input" not in text.lower() or "Remove-Item" not in text
        assert "results" not in text.lower() or "Remove-Item" not in text
    update = Path("persistent_root/UPDATE_HARNESS.ps1").read_text("ascii")
    assert "updates" in update
    assert "version.json" in update
    assert ".venv_harness" in update
    assert "$ErrorActionPreference = 'Stop'" in update
    rollback = Path("persistent_root/ROLLBACK_HARNESS.ps1").read_text("ascii")
    assert "backup" in rollback
    assert "version.json" in rollback


def test_persistent_runtime_reuses_ready_environment_and_archives_applied_update():
    current_runner = Path("RUN_REMOTE_WORD_HARNESS.ps1").read_text("ascii")
    assert ".word_replica_ready" in current_runner
    assert "if (-not (Test-Path $ReadyMarker))" in current_runner
    updater = Path("persistent_root/UPDATE_HARNESS.ps1").read_text("ascii")
    assert "applied" in updater
    assert "Move-Item" in updater


def test_persistent_launcher_self_heals_required_data_directories():
    text = Path("persistent_root/RUN_PERSISTENT_HARNESS.ps1").read_text(encoding="ascii")
    for dirname in ("realworld_input", "results", "backup", "updates"):
        assert f"Join-Path $Root '{dirname}'" in text
    assert "New-Item -ItemType Directory" in text
    assert "-Force" in text


def test_update_launcher_self_heals_update_and_backup_directories():
    text = Path("persistent_root/UPDATE_HARNESS.ps1").read_text(encoding="ascii")
    assert "Join-Path $Root 'updates'" in text
    assert "Join-Path $Root 'backup'" in text
    assert "New-Item -ItemType Directory" in text
    assert "-Force" in text
