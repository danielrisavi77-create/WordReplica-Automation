from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "scripts" / "codex_automation" / "bootstrap_windows.ps1"


def _script_text() -> str:
    return SCRIPT.read_text(encoding="utf-8")


def test_bootstrap_seeds_private_repo_and_automation_dev_branch():
    text = _script_text()
    assert "danielrisavi77-create/WordReplica-Automation.git" in text
    assert "automation-dev" in text
    assert "Run-Git push origin main" in text
    assert "Run-Git push -u origin automation-dev" in text
    assert "baseline_main.zip" in text
    assert "automation_final.zip" in text


def test_bootstrap_keeps_golden_local_and_creates_editable_test_venv():
    text = _script_text()
    assert r"C:\WordReplica-Automation" in text
    assert "Glavna verzija rektorova (grupno)(1).docx" in text
    assert '.venv' in text
    assert 'pip install -e ".[test]"' in text
    assert "WORD_REPLICA_AUTOMATION_ROOT" in text
    assert "WORD_REPLICA_GOLDEN_DOCX" in text


def test_bootstrap_never_uploads_artifacts_or_globally_kills_word():
    text = _script_text().lower()
    assert "upload-artifact" not in text
    assert "taskkill /im winword" not in text
    assert "*.docx" not in text or "git" not in text.split("*.docx")[0][-40:]


def test_bootstrap_runs_regression_suite_before_declaring_ready():
    text = _script_text()
    assert "-m pytest -q" in text
    assert "AUTOMATION SETUP: PASS" in text


def test_bootstrap_exposes_golden_environment_before_automation_tests():
    text = _script_text()
    env_pos = text.index('$env:WORD_REPLICA_GOLDEN_DOCX')
    automation_pos = text.index('Creating automation-dev and applying Codex automation layer')
    assert env_pos < automation_pos


def test_bootstrap_tolerates_missing_local_git_identity():
    text = _script_text()
    assert '(& git config user.name).Trim()' not in text
    assert '(& git config user.email).Trim()' not in text
