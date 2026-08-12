from pathlib import Path

from scripts.codex_automation import process as process_module
from scripts.codex_automation import word_process as word_process_module
from scripts.codex_automation.process import ChildResult, run_owned_child
from scripts.codex_automation.word_process import OwnedWordProcess


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
    monkeypatch.setattr("scripts.codex_automation.process.process_tree_pids", lambda root_pid: {root_pid})
    monkeypatch.setattr(
        "scripts.codex_automation.process.kill_process_tree",
        lambda root_pid: fake.kill(),
    )
    monkeypatch.setattr("scripts.codex_automation.process.list_owned_word_processes", lambda path: ["owned"])
    monkeypatch.setattr(
        "scripts.codex_automation.process.terminate_owned_word_processes",
        lambda records, expected_owner_pids: calls.append((records, expected_owner_pids)) or [9001],
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
    assert calls == [(["owned"], {4321})]
    assert result.terminated_word_pids == [9001]


def test_timeout_accepts_verified_python_descendant_as_owned_word_owner(tmp_path, monkeypatch):
    fake = FakeProcess()
    tree_kills = []
    word_kills = []
    owned = OwnedWordProcess(
        pid=9001,
        hwnd=99,
        owner_process_pid=8765,
        role="interactive",
        started_filetime=123456789,
    )
    monkeypatch.setattr("scripts.codex_automation.process.subprocess.Popen", lambda *a, **k: fake)
    monkeypatch.setattr(
        "scripts.codex_automation.process.process_tree_pids",
        lambda root_pid: {root_pid, 8765},
        raising=False,
    )
    monkeypatch.setattr(
        "scripts.codex_automation.process.kill_process_tree",
        lambda root_pid: (tree_kills.append(root_pid), fake.kill()),
        raising=False,
    )
    monkeypatch.setattr("scripts.codex_automation.process.list_owned_word_processes", lambda path: [owned])
    monkeypatch.setattr(
        word_process_module,
        "_windows_image_name",
        lambda pid: "WINWORD.EXE",
    )
    monkeypatch.setattr(
        word_process_module,
        "_windows_process_creation_filetime",
        lambda pid: 123456789,
    )
    monkeypatch.setattr(
        word_process_module,
        "_windows_kill_pid",
        lambda pid: word_kills.append(pid),
    )

    result = run_owned_child(
        ["python", "child.py"],
        timeout_seconds=1,
        stdout_path=tmp_path / "out.log",
        stderr_path=tmp_path / "err.log",
        ownership_file=tmp_path / "owned.json",
    )

    assert result.terminated_word_pids == [9001]
    assert tree_kills == [4321]
    assert word_kills == [9001]


def test_descendant_process_ids_include_nested_children_but_not_unrelated_processes():
    parent_by_pid = {
        4321: 100,
        8765: 4321,
        9000: 8765,
        7777: 100,
    }

    assert process_module.descendant_process_pids(4321, parent_by_pid) == {4321, 8765, 9000}
