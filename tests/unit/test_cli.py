from pathlib import Path

from word_replica.cli import build_parser, run_repair_poc, run_repair_poc_resume


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
