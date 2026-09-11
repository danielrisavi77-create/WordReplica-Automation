from __future__ import annotations

import hashlib
import io
import json
import threading
from email.message import Message
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.request import Request

import pytest

from word_replica.domain.enums import RendererChoice
from word_replica.runner.http_transport import TransportError
from word_replica.runner.lekta_claim import LaunchTicket
from word_replica.runner.one_shot import OneShotRunner
from word_replica.runner.trust_store import prepare_release_trust_store


JOB_ID = "33333333-3333-4333-8333-333333333333"
TOKEN = "A" * 43
EXE_NAME = f"LektaRepair-{JOB_ID}-{TOKEN}.exe"


class FakeResponse:
    def __init__(self, body: bytes, *, content_length: int | None = None) -> None:
        self._stream = io.BytesIO(body)
        self.headers = Message()
        if content_length is not None:
            self.headers["Content-Length"] = str(content_length)

    def read(self, size: int = -1) -> bytes:
        return self._stream.read(size)

    def __enter__(self) -> "FakeResponse":
        return self

    def __exit__(self, *_args: object) -> None:
        return None


@pytest.mark.parametrize(
    ("source", "rewritten"),
    [
        ("https://127.0.0.1:8765/claim", "http://127.0.0.1:8765/claim"),
        ("https://[::1]:8765/status?q=1", "http://[::1]:8765/status?q=1"),
    ],
)
def test_development_transport_downgrades_only_https_form_exact_loopback(
    source: str, rewritten: str
) -> None:
    from word_replica.runner.development_entry import DevelopmentLoopbackTransport

    captured: list[Request] = []

    def open_request(request: Request, *, timeout: float) -> FakeResponse:
        captured.append(request)
        assert timeout == 3.5
        return FakeResponse(b'{"ok":true}', content_length=11)

    transport = DevelopmentLoopbackTransport(
        open_request=open_request, timeout_seconds=3.5
    )

    assert transport.post_json(source, {"jobId": JOB_ID}) == {"ok": True}
    assert captured[0].full_url == rewritten
    assert json.loads(captured[0].data or b"") == {"jobId": JOB_ID}


@pytest.mark.parametrize(
    "url",
    [
        "http://127.0.0.1:8765/claim",
        "https://localhost:8765/claim",
        "https://127.0.0.2:8765/claim",
        "https://example.test/claim",
        "https://user@127.0.0.1:8765/claim",
        "https://user:secret@[::1]:8765/claim",
    ],
)
def test_development_transport_rejects_non_loopback_credentials_and_direct_http(
    url: str,
) -> None:
    from word_replica.runner.development_entry import DevelopmentLoopbackTransport

    called = False

    def open_request(_request: Request, *, timeout: float) -> FakeResponse:
        nonlocal called
        called = True
        return FakeResponse(b"{}")

    with pytest.raises(TransportError, match="exact loopback"):
        DevelopmentLoopbackTransport(open_request=open_request).post_json(url, {})
    assert called is False


def test_development_transport_rejects_redirects() -> None:
    from word_replica.runner.development_entry import DevelopmentLoopbackTransport

    requested_paths: list[str] = []

    class RedirectHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 - stdlib callback name
            requested_paths.append(self.path)
            if self.path == "/start":
                self.send_response(302)
                self.send_header(
                    "Location", f"https://127.0.0.1:{self.server.server_port}/target"
                )
                self.end_headers()
                return
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"redirect followed")

        def log_message(self, _format: str, *_args: object) -> None:
            return

    server = ThreadingHTTPServer(("127.0.0.1", 0), RedirectHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        with pytest.raises(TransportError, match="redirect"):
            DevelopmentLoopbackTransport().get_bytes(
                f"https://127.0.0.1:{server.server_port}/start",
                maximum_bytes=100,
            )
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)

    assert requested_paths == ["/start"]


def test_development_transport_ignores_environment_http_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from word_replica.runner.development_entry import DevelopmentLoopbackTransport

    origin_paths: list[str] = []
    proxy_paths: list[str] = []

    class OriginHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 - stdlib callback name
            origin_paths.append(self.path)
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"origin")

        def log_message(self, _format: str, *_args: object) -> None:
            return

    class ProxyHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802 - stdlib callback name
            proxy_paths.append(self.path)
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"proxy")

        def log_message(self, _format: str, *_args: object) -> None:
            return

    origin = ThreadingHTTPServer(("127.0.0.1", 0), OriginHandler)
    proxy = ThreadingHTTPServer(("127.0.0.1", 0), ProxyHandler)
    origin_thread = threading.Thread(target=origin.serve_forever, daemon=True)
    proxy_thread = threading.Thread(target=proxy.serve_forever, daemon=True)
    origin_thread.start()
    proxy_thread.start()
    monkeypatch.setenv("HTTP_PROXY", f"http://127.0.0.1:{proxy.server_port}")
    monkeypatch.setenv("http_proxy", f"http://127.0.0.1:{proxy.server_port}")
    monkeypatch.delenv("NO_PROXY", raising=False)
    monkeypatch.delenv("no_proxy", raising=False)
    monkeypatch.setattr("urllib.request.proxy_bypass", lambda _host: False)
    try:
        body = DevelopmentLoopbackTransport().get_bytes(
            f"https://127.0.0.1:{origin.server_port}/source.docx",
            maximum_bytes=100,
        )
    finally:
        origin.shutdown()
        proxy.shutdown()
        origin.server_close()
        proxy.server_close()
        origin_thread.join(timeout=5)
        proxy_thread.join(timeout=5)

    assert body == b"origin"
    assert origin_paths == ["/source.docx"]
    assert proxy_paths == []


class _Result:
    def __init__(self, output_path: Path) -> None:
        self.output_path = output_path
        self.report = self

    def to_json(self) -> dict[str, object]:
        return {"zeta": 2, "message": "točan izvještaj"}


class _Runner:
    def __init__(self, *, interrupted: bool, output_path: Path) -> None:
        self.interrupted = interrupted
        self.output_path = output_path
        self.run_calls = []
        self.resume_calls = []

    def has_interrupted_job(self, job_id: str) -> bool:
        return self.interrupted and job_id == JOB_ID

    def run(self, ticket, config, *, on_claimed=None):
        self.run_calls.append((ticket, config, on_claimed))
        return _Result(self.output_path)

    def resume_interrupted(self, job_id: str, config):
        self.resume_calls.append((job_id, config))
        return _Result(self.output_path)


@pytest.mark.parametrize("interrupted", [False, True])
def test_development_entry_is_noninteractive_and_uses_the_production_filename_ticket(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, interrupted: bool
) -> None:
    from word_replica.runner import portable_entry
    from word_replica.runner.development_entry import run_development_e2e

    monkeypatch.setattr(
        portable_entry, "_choose_output",
        lambda: pytest.fail("development E2E opened the folder chooser"),
    )
    monkeypatch.setattr(
        portable_entry, "schedule_portable_cleanup",
        lambda _path: pytest.fail("development E2E scheduled self-deletion"),
    )
    output_dir = tmp_path / "output"
    state_dir = tmp_path / "state"
    runner = _Runner(interrupted=interrupted, output_path=output_dir / "fixed.docx")

    result = run_development_e2e(
        Path("C:/Downloads") / EXE_NAME,
        claim_endpoint="https://127.0.0.1:8765/claim",
        status_endpoint="https://[::1]:8766/status",
        output_dir=output_dir,
        state_dir=state_dir,
        runner=runner,
    )

    assert result == 0
    calls = runner.resume_calls if interrupted else runner.run_calls
    assert len(calls) == 1
    if interrupted:
        job_id, config = calls[0]
        assert job_id == JOB_ID
        assert runner.run_calls == []
    else:
        ticket, config, on_claimed = calls[0]
        assert ticket == LaunchTicket(job_id=JOB_ID, claim_token=TOKEN)
        assert on_claimed is None
        assert runner.resume_calls == []
    assert config.claim_endpoint == "https://127.0.0.1:8765/claim"
    assert config.status_endpoint == "https://[::1]:8766/status"
    assert config.output_dir == output_dir.resolve()


def test_development_runner_reuses_one_shot_and_pure_docx_service(tmp_path: Path) -> None:
    from word_replica.runner.development_entry import create_development_runner

    public_key = (
        Path(__file__).resolve().parents[1]
        / "fixtures"
        / "repair_contract_v1"
        / "public-key.spki.b64url"
    )
    trust_store = tmp_path / "trusted_keys.json"
    prepare_release_trust_store(
        public_key_path=public_key,
        key_id="lekta-dev-test",
        destination=trust_store,
    )

    runner = create_development_runner(
        state_dir=tmp_path / "state", trust_path=trust_store
    )

    assert isinstance(runner, OneShotRunner)
    assert runner.retry_store.root == tmp_path / "state"
    assert runner.package_service.rebuild_options.renderer is RendererChoice.DOCX


def test_development_arguments_require_the_explicit_e2e_flag(tmp_path: Path) -> None:
    from word_replica.runner.development_entry import main

    runner = _Runner(interrupted=False, output_path=tmp_path / "output" / "fixed.docx")
    common = [
        "--claim-endpoint", "https://127.0.0.1:8765/claim",
        "--status-endpoint", "https://127.0.0.1:8765/status",
        "--output-directory", str(tmp_path / "output"),
        "--state-directory", str(tmp_path / "state"),
    ]

    assert main(common, executable_path=Path(EXE_NAME), runner=runner) == 2
    assert runner.run_calls == []
    assert main(
        ["--development-e2e", *common],
        executable_path=Path(EXE_NAME),
        runner=runner,
    ) == 0
    assert len(runner.run_calls) == 1


def test_development_entry_writes_sanitized_failure_diagnostic(tmp_path: Path) -> None:
    from word_replica.runner.development_entry import main

    private_key = "-----BEGIN PRIVATE KEY-----\nprivate-key-material\n-----END PRIVATE KEY-----"
    diagnostic_path = tmp_path / "development-error.json"
    portable_name = EXE_NAME

    class FailingRunner(_Runner):
        def run(self, ticket, config, *, on_claimed=None):
            raise RuntimeError(
                f"failure for {portable_name} job={JOB_ID} token={TOKEN} {private_key}"
            )

    runner = FailingRunner(interrupted=False, output_path=tmp_path / "unused.docx")
    exit_code = main(
        [
            "--development-e2e",
            "--claim-endpoint", "https://127.0.0.1:8765/claim",
            "--status-endpoint", "https://127.0.0.1:8765/status",
            "--output-directory", str(tmp_path / "output"),
            "--state-directory", str(tmp_path / "state"),
            "--diagnostic-path", str(diagnostic_path),
        ],
        executable_path=Path(portable_name),
        runner=runner,
    )

    assert exit_code == 2
    diagnostic = json.loads(diagnostic_path.read_text(encoding="utf-8"))
    assert diagnostic["error"]["type"] == "RuntimeError"
    assert "failure for" in diagnostic["error"]["message"]
    assert "RuntimeError" in diagnostic["error"]["traceback"]
    rendered = json.dumps(diagnostic)
    assert portable_name not in rendered
    assert JOB_ID not in rendered
    assert TOKEN not in rendered
    assert private_key not in rendered


def test_development_entry_persists_exact_canonical_signed_report(tmp_path: Path) -> None:
    from word_replica.runner.development_entry import main
    from word_replica.runner.one_shot import _report_sha256

    report_path = tmp_path / "signed-report.json"
    runner = _Runner(
        interrupted=False, output_path=tmp_path / "output" / "fixed.docx"
    )

    exit_code = main(
        [
            "--development-e2e",
            "--claim-endpoint", "https://127.0.0.1:8765/claim",
            "--status-endpoint", "https://127.0.0.1:8765/status",
            "--output-directory", str(tmp_path / "output"),
            "--state-directory", str(tmp_path / "state"),
            "--report-path", str(report_path),
        ],
        executable_path=Path(EXE_NAME),
        runner=runner,
    )

    assert exit_code == 0
    expected_report = _Result(tmp_path / "unused.docx")
    report_bytes = report_path.read_bytes()
    assert report_bytes == '{"message":"točan izvještaj","zeta":2}'.encode()
    assert hashlib.sha256(report_bytes).hexdigest() == _report_sha256(expected_report)
