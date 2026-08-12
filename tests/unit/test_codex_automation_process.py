from pathlib import Path

from scripts.codex_automation.process import ChildResult, run_owned_child


class FakeProcess:
    def __init__(self):
        self.pid = 4321
        self.returncode = None
        self.killed = False
    def wait(self, timeout=None):
        import subprocess
        raise subprocess.TimeoutExpired(cmd="x", timeout=timeout)
    def kill(self):
        self.killed = True
        self.returncode = -9


def test_timeout_kills_child_and_only_owned_word_records(tmp_path, monkeypatch):
    fake = FakeProcess()
    calls = []
    monkeypatch.setattr("scripts.codex_automation.process.subprocess.Popen", lambda *a, **k: fake)
    monkeypatch.setattr("scripts.codex_automation.process.list_owned_word_processes", lambda path: ["owned"])
    monkeypatch.setattr(
        "scripts.codex_automation.process.terminate_owned_word_processes",
        lambda records, expected_owner_pid: calls.append((records, expected_owner_pid)) or [9001],
    )

    result = run_owned_child(
        ["python", "child.py"],
        timeout_seconds=1,
        stdout_path=tmp_path / "out.log",
        stderr_path=tmp_path / "err.log",
        ownership_file=tmp_path / "owned.json",
    )

    assert isinstance(result, ChildResult)
    assert result.timed_out is True
    assert fake.killed is True
    assert calls == [(["owned"], 4321)]
    assert result.terminated_word_pids == [9001]
