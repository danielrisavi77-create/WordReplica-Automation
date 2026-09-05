import json
from pathlib import Path

from word_replica.renderers.word_ownership import record_owned_word, clear_owned_word


class FakeApplication:
    Hwnd = 123


class FakeApplicationWithoutHwnd:
    @property
    def Hwnd(self):
        raise AttributeError("Word.Application.Hwnd")


def test_record_owned_word_writes_exact_pid_owner_and_role(tmp_path, monkeypatch):
    record = tmp_path / "owned.json"
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNERSHIP_FILE", str(record))

    pid = record_owned_word(FakeApplication(), role="interactive", pid_resolver=lambda hwnd: 4567, process_identity_resolver=lambda pid: 987654321, owner_pid=111)

    assert pid == 4567
    payload = json.loads(record.read_text(encoding="utf-8"))
    assert payload["pid"] == 4567
    assert payload["hwnd"] == 123
    assert payload["owner_process_pid"] == 111
    assert payload["role"] == "interactive"
    assert payload["started_filetime"] == 987654321

    clear_owned_word(4567)
    assert not record.exists()


def test_record_owned_word_uses_new_word_process_when_application_has_no_hwnd(tmp_path, monkeypatch):
    record = tmp_path / "owned.json"
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNERSHIP_FILE", str(record))

    pid = record_owned_word(
        FakeApplicationWithoutHwnd(),
        role="interactive",
        existing_word_pids={1001},
        word_process_pids_resolver=lambda: {1001, 4567},
        process_identity_resolver=lambda process_pid: 987654321,
        owner_pid=111,
    )

    assert pid == 4567
    payload = json.loads(record.read_text(encoding="utf-8"))
    assert payload["pid"] == 4567
    assert payload["hwnd"] == 0
    assert payload["owner_process_pid"] == 111


def test_clear_does_not_remove_record_for_different_pid(tmp_path, monkeypatch):
    record = tmp_path / "owned.json"
    record.write_text(json.dumps({"pid": 9, "hwnd": 1, "owner_process_pid": 2, "role": "x"}), encoding="utf-8")
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNERSHIP_FILE", str(record))

    clear_owned_word(10)

    assert record.exists()


def test_multiple_owned_word_processes_coexist_and_clear_individually(tmp_path, monkeypatch):
    from word_replica.renderers.word_ownership import list_owned_words

    record = tmp_path / "owned.json"
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNERSHIP_FILE", str(record))

    record_owned_word(type("A", (), {"Hwnd": 1})(), role="interactive", pid_resolver=lambda hwnd: 1001, process_identity_resolver=lambda pid: 111111, owner_pid=77)
    record_owned_word(type("B", (), {"Hwnd": 2})(), role="pdf-export", pid_resolver=lambda hwnd: 1002, process_identity_resolver=lambda pid: 222222, owner_pid=77)

    assert [item["pid"] for item in list_owned_words()] == [1001, 1002]
    clear_owned_word(1002)
    assert [item["pid"] for item in list_owned_words()] == [1001]
    clear_owned_word(1001)
    assert not record.exists()


def test_record_owned_word_uses_inherited_orchestrator_owner_pid(tmp_path, monkeypatch):
    record = tmp_path / "owned.json"
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNERSHIP_FILE", str(record))
    monkeypatch.setenv("WORD_REPLICA_WORD_OWNER_PID", "2468")

    record_owned_word(
        FakeApplication(),
        role="interactive",
        pid_resolver=lambda hwnd: 4567,
        process_identity_resolver=lambda pid: 987654321,
    )

    payload = json.loads(record.read_text(encoding="utf-8"))
    assert payload["owner_process_pid"] == 2468
