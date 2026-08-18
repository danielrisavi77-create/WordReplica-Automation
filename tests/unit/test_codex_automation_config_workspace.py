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
