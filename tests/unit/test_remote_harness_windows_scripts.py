from pathlib import Path


def test_windows_harness_scripts_are_ascii_and_fail_fast():
    for filename in ("RUN_REMOTE_WORD_HARNESS.ps1","START_REMOTE_HARNESS.cmd"):
        raw=Path(filename).read_bytes()
        assert raw
        assert all(byte < 128 for byte in raw), filename
    text=Path("RUN_REMOTE_WORD_HARNESS.ps1").read_text(encoding="ascii")
    assert "$ErrorActionPreference = 'Stop'" in text
    assert "realworld_input" in text
    assert ".venv_harness" in text
    assert "--no-cache-dir" in text
    assert "REMOTE WORD HARNESS" in text
