import json
from pathlib import Path

from word_replica import __version__
from word_replica.cli import (
    REPAIR_ENGINE_VERSION,
    build_parser,
    run_repair_poc,
    run_repair_poc_resume,
    run_repair_runner,
)
from word_replica.runner.lekta_claim import LaunchTicket


def test_repair_cli_advertises_the_installed_wordreplica_engine_version():
    assert REPAIR_ENGINE_VERSION == __version__ == "0.1.0"


def test_rebuild_command_maps_modes():
    args = build_parser().parse_args([
        'rebuild', 'paper.docx', '--renderer', 'auto', '--visibility', 'visible',
        '--fidelity', 'full', '--metadata', 'fresh'
    ])
    assert args.command == 'rebuild'
    assert args.visibility == 'visible'
    assert args.fidelity == 'full'


def test_repair_poc_command_parses_only_explicit_paths():
    args = build_parser().parse_args([
        "repair-poc", "--original", "original.docx", "--target", "target.docx",
        "--contract", "contract.json", "--public-key", "public.spki",
        "--output-dir", "out",
    ])
    assert args.command == "repair-poc"
    assert args.original == Path("original.docx")
    assert args.target == Path("target.docx")
    assert args.contract == Path("contract.json")
    assert args.public_key == Path("public.spki")
    assert args.output_dir == Path("out")
    assert args.stop_after_event is None
    assert args.renderer == "word"


def test_repair_poc_accepts_pure_docx_renderer_for_licence_free_reconstruction():
    args = build_parser().parse_args([
        "repair-poc", "--original", "o.docx", "--target", "t.docx",
        "--contract", "c.json", "--public-key", "k.spki", "--output-dir", "out",
        "--renderer", "pure-docx",
    ])
    assert args.renderer == "pure-docx"


def test_repair_poc_resume_command_requires_job_id():
    args = build_parser().parse_args([
        "repair-poc-resume", "--original", "o.docx", "--target", "t.docx",
        "--contract", "c.json", "--public-key", "k.spki", "--output-dir", "out",
        "--job-id", "11111111-1111-4111-8111-111111111111",
    ])
    assert args.command == "repair-poc-resume"
    assert args.job_id == "11111111-1111-4111-8111-111111111111"


def test_repair_runner_parses_only_launch_and_output_paths():
    args = build_parser().parse_args([
        "repair-runner", "--launch", "launch.json", "--output-dir", "out",
    ])
    assert args.command == "repair-runner"
    assert args.launch == Path("launch.json")
    assert args.output_dir == Path("out")
    assert not hasattr(args, "claim_endpoint")
    assert not hasattr(args, "public_key")


class FakeArgs:
    def __init__(self, output_dir, **overrides):
        self.original = Path("o.docx")
        self.target = Path("t.docx")
        self.contract = Path("c.json")
        self.public_key = Path("k.spki")
        self.output_dir = Path(output_dir)
        self.stop_after_event = None
        self.job_id = "job-1"
        for key, value in overrides.items():
            setattr(self, key, value)


class FakeReport:
    def __init__(self, *, status, full_pass, output_path=None):
        self.status = status
        self.full_pass = full_pass
        self.output_path = output_path

    def to_json(self):
        return {"status": self.status, "full_pass": self.full_pass}


class FakeService:
    def __init__(self, report):
        self._report = report
        self.run_calls = []
        self.resume_calls = []

    def run(self, request, **kwargs):
        self.run_calls.append((request, kwargs))
        return self._report

    def resume(self, request, job_id, **kwargs):
        self.resume_calls.append((request, job_id, kwargs))
        return self._report


def test_repair_poc_exits_0_and_hands_off_only_on_full_pass(tmp_path):
    report = FakeReport(status="FULL_PASS", full_pass=True, output_path=str(tmp_path / "out.docx"))
    service = FakeService(report)
    handoff_calls = []
    exit_code = run_repair_poc(FakeArgs(tmp_path), service=service, handoff=handoff_calls.append)
    assert exit_code == 0
    assert handoff_calls == [Path(report.output_path)]
    assert (tmp_path / "repair_poc_report.json").is_file()


def test_repair_poc_exits_1_and_never_hands_off_on_retryable(tmp_path):
    report = FakeReport(status="RETRYABLE", full_pass=False)
    service = FakeService(report)
    handoff_calls = []
    exit_code = run_repair_poc(FakeArgs(tmp_path), service=service, handoff=handoff_calls.append)
    assert exit_code == 1
    assert handoff_calls == []


def test_repair_poc_exits_2_and_never_hands_off_on_failed(tmp_path):
    report = FakeReport(status="FAILED", full_pass=False)
    service = FakeService(report)
    handoff_calls = []
    exit_code = run_repair_poc(FakeArgs(tmp_path), service=service, handoff=handoff_calls.append)
    assert exit_code == 2
    assert handoff_calls == []


def test_repair_poc_exits_2_on_invalid_package_without_calling_handoff(tmp_path):
    from word_replica.domain.errors import RepairPackageError

    class RaisingService:
        def run(self, request, **kwargs):
            raise RepairPackageError("source-hash-mismatch")

    handoff_calls = []
    exit_code = run_repair_poc(FakeArgs(tmp_path), service=RaisingService(), handoff=handoff_calls.append)
    assert exit_code == 2
    assert handoff_calls == []


def test_repair_poc_resume_calls_service_resume_with_the_given_job_id(tmp_path):
    report = FakeReport(status="FULL_PASS", full_pass=True, output_path=str(tmp_path / "out.docx"))
    service = FakeService(report)
    handoff_calls = []
    exit_code = run_repair_poc_resume(
        FakeArgs(tmp_path, job_id="job-42"), service=service, handoff=handoff_calls.append
    )
    assert exit_code == 0
    assert service.resume_calls[0][1] == "job-42"


def test_stop_after_event_builds_a_control_that_stops_only_at_the_boundary():
    from word_replica.cli import _stop_after_event_control
    from word_replica.domain.enums import InteractiveRunState

    control, observer = _stop_after_event_control(2)
    assert control is not None
    observer.event_completed(0, None)
    assert control.state is InteractiveRunState.RUNNING
    observer.event_completed(2, None)
    assert control.state is InteractiveRunState.STOPPED


def test_stop_after_event_is_a_noop_when_not_requested():
    from word_replica.cli import _stop_after_event_control

    control, observer = _stop_after_event_control(None)
    assert control is None
    assert observer is None


def test_repair_runner_replaces_claim_token_after_durable_claim_and_hands_off(tmp_path):
    launch_path = tmp_path / "launch.json"
    launch_path.write_text(json.dumps({
        "version": 1,
        "jobId": "33333333-3333-4333-8333-333333333333",
        "claimToken": "A" * 43,
    }), encoding="utf-8")
    output_path = tmp_path / "out" / "Seminar-popravljeno.docx"
    output_path.parent.mkdir()
    output_path.write_bytes(b"finished")
    calls = []

    class FakeRunner:
        def run(self, ticket, config, *, on_claimed):
            calls.append((ticket, config))
            on_claimed(ticket.job_id)
            resume_pointer = json.loads(launch_path.read_text(encoding="utf-8"))
            assert resume_pointer == {
                "version": 1,
                "jobId": ticket.job_id,
                "resume": True,
            }
            return type("Result", (), {"output_path": output_path})()

    handoffs = []
    args = type("Args", (), {"launch": launch_path, "output_dir": output_path.parent})()
    assert run_repair_runner(args, runner=FakeRunner(), handoff=handoffs.append) == 0

    ticket, config = calls[0]
    assert isinstance(ticket, LaunchTicket)
    assert ticket.job_id == "33333333-3333-4333-8333-333333333333"
    assert config.output_dir == output_path.parent.resolve()
    assert config.status_endpoint == (
        "https://zrrjttizjyfcxmcpgzml.supabase.co/functions/v1/repair-local-status"
    )
    assert handoffs == [output_path]
    assert not launch_path.exists()


def test_repair_runner_keeps_unconsumed_launch_on_invalid_ticket(tmp_path):
    launch_path = tmp_path / "launch.json"
    launch_path.write_text('{"claimToken":"secret"}', encoding="utf-8")
    called = False

    class FakeRunner:
        def run(self, *_args):
            nonlocal called
            called = True

    args = type("Args", (), {"launch": launch_path, "output_dir": tmp_path / "out"})()
    assert run_repair_runner(args, runner=FakeRunner(), handoff=lambda _path: None) == 2
    assert called is False
    assert launch_path.exists()


def test_repair_runner_resumes_pending_receipt_without_reusing_claim_token(tmp_path):
    launch_path = tmp_path / "launch.json"
    launch_path.write_text(json.dumps({
        "version": 1,
        "jobId": "33333333-3333-4333-8333-333333333333",
        "claimToken": "A" * 43,
    }), encoding="utf-8")
    output_path = tmp_path / "out" / "Seminar-popravljeno.docx"
    output_path.parent.mkdir()
    output_path.write_bytes(b"finished")

    class FakeRunner:
        def __init__(self):
            self.run_calls = 0
            self.resume_calls = 0

        def run(self, ticket, config, *, on_claimed):
            self.run_calls += 1
            on_claimed(ticket.job_id)
            raise OSError("offline after local FULL_PASS")

        def resume_interrupted(self, job_id, config):
            self.resume_calls += 1
            assert job_id == "33333333-3333-4333-8333-333333333333"
            return type("Result", (), {"output_path": output_path})()

    runner = FakeRunner()
    handoffs = []
    args = type("Args", (), {"launch": launch_path, "output_dir": output_path.parent})()

    assert run_repair_runner(args, runner=runner, handoff=handoffs.append) == 2
    resume_pointer = json.loads(launch_path.read_text(encoding="utf-8"))
    assert resume_pointer == {
        "version": 1,
        "jobId": "33333333-3333-4333-8333-333333333333",
        "resume": True,
    }
    assert "claimToken" not in launch_path.read_text(encoding="utf-8")

    assert run_repair_runner(args, runner=runner, handoff=handoffs.append) == 0
    assert runner.run_calls == 1
    assert runner.resume_calls == 1
    assert handoffs == [output_path]
    assert not launch_path.exists()


def test_repair_runner_resumes_dpapi_job_even_when_original_ticket_survived_crash(tmp_path):
    job_id = "33333333-3333-4333-8333-333333333333"
    launch_path = tmp_path / "launch.json"
    launch_path.write_text(json.dumps({
        "version": 1,
        "jobId": job_id,
        "claimToken": "A" * 43,
    }), encoding="utf-8")
    output_path = tmp_path / "out" / "Seminar-popravljeno.docx"
    output_path.parent.mkdir()
    output_path.write_bytes(b"resumed")

    class FakeRunner:
        def __init__(self):
            self.run_calls = 0
            self.resume_calls = 0

        def has_interrupted_job(self, candidate_job_id):
            assert candidate_job_id == job_id
            return True

        def run(self, *_args, **_kwargs):
            self.run_calls += 1
            raise AssertionError("consumed claim token must not be submitted again")

        def resume_interrupted(self, candidate_job_id, config):
            self.resume_calls += 1
            assert candidate_job_id == job_id
            scrubbed = json.loads(launch_path.read_text(encoding="utf-8"))
            assert scrubbed == {"version": 1, "jobId": job_id, "resume": True}
            return type("Result", (), {"output_path": output_path})()

    runner = FakeRunner()
    handoffs = []
    args = type("Args", (), {"launch": launch_path, "output_dir": output_path.parent})()

    assert run_repair_runner(args, runner=runner, handoff=handoffs.append) == 0
    assert runner.run_calls == 0
    assert runner.resume_calls == 1
    assert handoffs == [output_path]
    assert not launch_path.exists()
