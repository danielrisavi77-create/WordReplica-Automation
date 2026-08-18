import json
from pathlib import Path

from scripts.codex_automation.word_process import OwnedWordProcess, read_owned_word_process, may_terminate_owned_word


def test_owned_word_record_requires_matching_child_pid_and_winword(tmp_path):
    path = tmp_path / "owned_word.json"
    path.write_text(json.dumps({
        "pid": 4242,
        "hwnd": 99,
        "owner_process_pid": 111,
        "role": "interactive",
        "started_filetime": 123456789,
    }), encoding="utf-8")

    record = read_owned_word_process(path)
    assert record == OwnedWordProcess(pid=4242, hwnd=99, owner_process_pid=111, role="interactive", started_filetime=123456789)
    assert may_terminate_owned_word(record, expected_owner_pid=111, image_name="WINWORD.EXE", live_started_filetime=123456789) is True
    assert may_terminate_owned_word(record, expected_owner_pid=112, image_name="WINWORD.EXE", live_started_filetime=123456789) is False
    assert may_terminate_owned_word(record, expected_owner_pid=111, image_name="notepad.exe", live_started_filetime=123456789) is False
    assert may_terminate_owned_word(record, expected_owner_pid=111, image_name="WINWORD.EXE", live_started_filetime=999999999) is False


def test_registry_reader_and_terminator_only_targets_verified_owned_winword(tmp_path):
    from scripts.codex_automation.word_process import list_owned_word_processes, terminate_owned_word_processes

    path = tmp_path / "owned_word.json"
    path.write_text(json.dumps({
        "schema_version": 1,
        "processes": [
            {"pid": 1001, "hwnd": 1, "owner_process_pid": 111, "role": "interactive", "started_filetime": 10001},
            {"pid": 1002, "hwnd": 2, "owner_process_pid": 111, "role": "pdf-export", "started_filetime": 10002},
            {"pid": 2001, "hwnd": 3, "owner_process_pid": 999, "role": "other", "started_filetime": 20001},
        ],
    }), encoding="utf-8")

    records = list_owned_word_processes(path)
    killed = []
    result = terminate_owned_word_processes(
        records,
        expected_owner_pid=111,
        image_resolver=lambda pid: "WINWORD.EXE" if pid in {1001, 2001} else "notepad.exe",
        identity_resolver=lambda pid: {1001: 10001, 1002: 10002, 2001: 20001}[pid],
        killer=lambda pid: killed.append(pid),
    )

    assert killed == [1001]
    assert result == [1001]


def test_terminator_refuses_reused_pid_even_if_image_and_owner_match():
    from scripts.codex_automation.word_process import terminate_owned_word_processes

    record = OwnedWordProcess(
        pid=4242, hwnd=9, owner_process_pid=111, role="interactive", started_filetime=123456789
    )
    killed = []
    result = terminate_owned_word_processes(
        [record],
        expected_owner_pid=111,
        image_resolver=lambda pid: "WINWORD.EXE",
        identity_resolver=lambda pid: 987654321,
        killer=lambda pid: killed.append(pid),
    )
    assert result == []
    assert killed == []
