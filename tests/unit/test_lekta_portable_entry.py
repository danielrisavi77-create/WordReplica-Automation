from pathlib import Path

import pytest

from word_replica.cli import execute_repair_runner_ticket
from word_replica.runner.lekta_claim import ClaimProtocolError, LaunchTicket
from word_replica.runner import portable_entry
from word_replica.runner.portable_entry import (
    parse_portable_executable_name,
    run_portable_self_test,
    run_portable_entry,
    schedule_portable_cleanup,
)
from word_replica.runner.trust_store import prepare_release_trust_store


JOB_ID = "33333333-3333-4333-8333-333333333333"
TOKEN = "A" * 43
EXE_NAME = f"LektaRepair-{JOB_ID}-{TOKEN}.exe"


def test_frozen_self_test_loads_only_the_expected_packaged_contract_key(tmp_path):
    public_key = (
        Path(__file__).resolve().parents[1]
        / "fixtures"
        / "repair_contract_v1"
        / "public-key.spki.b64url"
    )
    trust_store = tmp_path / "trusted_keys.json"
    prepare_release_trust_store(
        public_key_path=public_key,
        key_id="lekta-release-test",
        destination=trust_store,
    )

    assert run_portable_self_test(
        "lekta-release-test", trust_path=trust_store
    ) == 0
    assert run_portable_self_test(
        "unexpected-key", trust_path=trust_store
    ) == 2


def test_portable_executable_name_is_the_only_launch_envelope():
    ticket = parse_portable_executable_name(Path("C:/Users/Test/Downloads") / EXE_NAME)

    assert ticket == LaunchTicket(job_id=JOB_ID, claim_token=TOKEN)


@pytest.mark.parametrize(
    "name",
    [
        "LektaRepair.exe",
        f"LektaRepair-{JOB_ID}-short.exe",
        f"WordReplica-{JOB_ID}-{TOKEN}.exe",
        f"LektaRepair-{JOB_ID}-{TOKEN}.exe.json",
        f"LektaRepair-../../-{TOKEN}.exe",
    ],
)
def test_portable_executable_name_fails_closed(name):
    with pytest.raises(ClaimProtocolError):
        parse_portable_executable_name(Path(name))


def test_portable_entry_selects_output_before_claim_and_passes_ticket_in_memory(tmp_path):
    order = []
    calls = []
    cleanup = []

    def choose_output():
        order.append("choose")
        return tmp_path

    def execute(ticket, output_dir):
        order.append("execute")
        calls.append((ticket, output_dir))
        return 0

    code = run_portable_entry(
        Path("C:/Users/Test/Downloads") / EXE_NAME,
        choose_output=choose_output,
        execute=execute,
        show_error=lambda _message: pytest.fail("unexpected error"),
        schedule_cleanup=cleanup.append,
    )

    assert code == 0
    assert order == ["choose", "execute"]
    assert calls == [(LaunchTicket(job_id=JOB_ID, claim_token=TOKEN), tmp_path)]
    assert cleanup == [Path("C:/Users/Test/Downloads") / EXE_NAME]
    assert list(tmp_path.iterdir()) == []


def test_portable_entry_cancel_does_not_consume_claim():
    executed = False
    cleanup = []

    def execute(_ticket, _output_dir):
        nonlocal executed
        executed = True
        return 0

    code = run_portable_entry(
        Path(EXE_NAME),
        choose_output=lambda: None,
        execute=execute,
        show_error=lambda _message: pytest.fail("unexpected error"),
        schedule_cleanup=cleanup.append,
    )

    assert code == 1
    assert executed is False
    assert cleanup == []


def test_portable_entry_failure_keeps_exe_for_same_device_retry(tmp_path):
    cleanup = []

    code = run_portable_entry(
        Path(EXE_NAME),
        choose_output=lambda: tmp_path,
        execute=lambda _ticket, _output: 2,
        show_error=lambda _message: pytest.fail("unexpected error"),
        schedule_cleanup=cleanup.append,
    )

    assert code == 2
    assert cleanup == []



def test_self_cleanup_uses_hidden_ps51_compatible_exact_target(tmp_path, monkeypatch):
    executable = tmp_path / EXE_NAME
    executable.write_bytes(b"signed-runner")
    monkeypatch.setattr(portable_entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(portable_entry.sys, "executable", str(executable))
    calls = []

    def spawn(arguments, **kwargs):
        calls.append((arguments, kwargs))
        return object()

    assert schedule_portable_cleanup(
        executable,
        spawn=spawn,
        parent_pid=4321,
    ) is True

    assert len(calls) == 1
    arguments, kwargs = calls[0]
    assert arguments[:6] == [
        "powershell.exe", "-NoProfile", "-NonInteractive",
        "-WindowStyle", "Hidden", "-Command",
    ]
    script = arguments[6]
    assert "Wait-Process" not in script
    assert "Get-Process" in script
    assert str(executable) not in script
    assert kwargs["env"]["WORDREPLICA_DELETE_TARGET"] == str(executable.resolve())
    assert kwargs["env"]["WORDREPLICA_PARENT_PID"] == "4321"
    assert kwargs["creationflags"] != 0
    assert kwargs["close_fds"] is True


class _Result:
    def __init__(self, output_path):
        self.output_path = output_path


class _Runner:
    def __init__(self, *, interrupted, output_path):
        self.interrupted = interrupted
        self.output_path = output_path
        self.run_calls = []
        self.resume_calls = []

    def has_interrupted_job(self, job_id):
        return self.interrupted and job_id == JOB_ID

    def run(self, ticket, config, *, on_claimed=None):
        self.run_calls.append((ticket, config, on_claimed))
        return _Result(self.output_path)

    def resume_interrupted(self, job_id, config):
        self.resume_calls.append((job_id, config))
        return _Result(self.output_path)


@pytest.mark.parametrize("interrupted", [False, True])
def test_in_memory_ticket_uses_existing_dpapi_resume_before_any_second_claim(tmp_path, interrupted):
    output = tmp_path / "Seminar-popravljeno.docx"
    runner = _Runner(interrupted=interrupted, output_path=output)
    handed_off = []
    ticket = LaunchTicket(job_id=JOB_ID, claim_token=TOKEN)

    result = execute_repair_runner_ticket(
        ticket,
        tmp_path,
        runner=runner,
        handoff=handed_off.append,
    )

    assert result == 0
    assert handed_off == [output]
    if interrupted:
        assert runner.run_calls == []
        assert len(runner.resume_calls) == 1
        assert runner.resume_calls[0][0] == JOB_ID
    else:
        assert runner.resume_calls == []
        assert len(runner.run_calls) == 1
        assert runner.run_calls[0][0] == ticket
        assert runner.run_calls[0][2] is None
