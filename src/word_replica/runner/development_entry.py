"""Explicit, noninteractive entry for the unsigned development E2E runner.

Production endpoint validation stays HTTPS-only. Development E2E URLs are
therefore supplied in HTTPS form and translated to HTTP only after proving the
host is exactly an IP loopback address. The transport never follows redirects.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import secrets
import sys
import traceback
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit, urlunsplit
from urllib.request import HTTPRedirectHandler, ProxyHandler, Request, build_opener

from word_replica.runner.http_transport import (
    MAX_JSON_RESPONSE_BYTES,
    TransportError,
    _read_bounded,
)
from word_replica.runner.one_shot import OneShotRunner, OneShotRunnerConfig


_EXACT_LOOPBACK_HOSTS = frozenset({"127.0.0.1", "::1"})
_PORTABLE_FILENAME = re.compile(
    r"LektaRepair-[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}-[A-Za-z0-9_-]{43}\.exe",
    re.IGNORECASE,
)
_UUID = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}", re.IGNORECASE)
_TOKEN = re.compile(r"(?<![A-Za-z0-9_-])[A-Za-z0-9_-]{43}(?![A-Za-z0-9_-])")
_PRIVATE_KEY = re.compile(r"-----BEGIN [^-]*PRIVATE KEY-----.*?-----END [^-]*PRIVATE KEY-----", re.DOTALL)


def _development_http_url(url: str) -> str:
    try:
        parsed = urlsplit(url)
        port = parsed.port
    except (TypeError, ValueError) as exc:
        raise TransportError("development URL must use HTTPS-form exact loopback") from exc
    if (
        parsed.scheme != "https"
        or parsed.hostname not in _EXACT_LOOPBACK_HOSTS
        or parsed.username is not None
        or parsed.password is not None
        or parsed.fragment
    ):
        raise TransportError("development URL must use HTTPS-form exact loopback")
    host = f"[{parsed.hostname}]" if parsed.hostname == "::1" else parsed.hostname
    netloc = f"{host}:{port}" if port is not None else host
    return urlunsplit(("http", netloc, parsed.path, parsed.query, ""))


class _RejectRedirectHandler(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise TransportError("development loopback redirects are not allowed")


def _default_open_request(request: Request, *, timeout: float):
    return build_opener(ProxyHandler({}), _RejectRedirectHandler()).open(
        request, timeout=timeout
    )


@dataclass(slots=True)
class DevelopmentLoopbackTransport:
    """HTTP transport reachable only through HTTPS-form exact-loopback URLs."""

    open_request: Callable[..., Any] = _default_open_request
    timeout_seconds: float = 15.0

    def post_json(self, url: str, payload: dict) -> dict:
        request = Request(
            _development_http_url(url),
            data=json.dumps(
                payload, ensure_ascii=False, separators=(",", ":")
            ).encode("utf-8"),
            headers={"Accept": "application/json", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with self.open_request(request, timeout=self.timeout_seconds) as response:
                body = _read_bounded(response, MAX_JSON_RESPONSE_BYTES)
        except TransportError:
            raise
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            raise TransportError("development loopback request failed") from exc
        try:
            value = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise TransportError("invalid JSON response") from exc
        if not isinstance(value, dict):
            raise TransportError("JSON response must be an object")
        return value

    def get_bytes(self, url: str, *, maximum_bytes: int) -> bytes:
        request = Request(
            _development_http_url(url),
            headers={"Accept": "application/octet-stream"},
            method="GET",
        )
        try:
            with self.open_request(request, timeout=self.timeout_seconds) as response:
                return _read_bounded(response, maximum_bytes)
        except TransportError:
            raise
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            raise TransportError("development loopback request failed") from exc


def create_development_runner(
    *, state_dir: Path, trust_path: Path | None = None
) -> OneShotRunner:
    """Build the normal runner with only the development loopback transport."""
    from word_replica.cli import _default_repair_package_service
    from word_replica.runner.secure_retry import SecureRetryStore
    from word_replica.runner.trust_store import load_trust_keys

    packaged_trust_path = (
        Path(trust_path)
        if trust_path is not None
        else Path(__file__).resolve().parent / "trusted_keys.json"
    )
    return OneShotRunner(
        preflight=lambda: None,
        transport=DevelopmentLoopbackTransport(),
        package_service=_default_repair_package_service(
            "pure-docx", visible_preview=False
        ),
        trust_keys=load_trust_keys(packaged_trust_path),
        retry_store=SecureRetryStore(Path(state_dir)),
    )


def run_development_e2e(
    executable_path: Path,
    *,
    claim_endpoint: str,
    status_endpoint: str,
    output_dir: Path,
    state_dir: Path,
    report_path: Path | None = None,
    runner: OneShotRunner | Any | None = None,
) -> int:
    """Run one development E2E ticket without GUI handoff or self-deletion."""
    from word_replica.runner.portable_entry import parse_portable_executable_name

    _development_http_url(claim_endpoint)
    _development_http_url(status_endpoint)
    ticket = parse_portable_executable_name(Path(executable_path))
    resolved_output = Path(output_dir).resolve()
    resolved_output.mkdir(parents=True, exist_ok=True)
    active_runner = runner or create_development_runner(
        state_dir=Path(state_dir).resolve()
    )
    config = OneShotRunnerConfig(
        claim_endpoint=claim_endpoint,
        status_endpoint=status_endpoint,
        output_dir=resolved_output,
    )
    if active_runner.has_interrupted_job(ticket.job_id):
        result = active_runner.resume_interrupted(ticket.job_id, config)
    else:
        result = active_runner.run(ticket, config, on_claimed=None)
    if report_path is not None:
        _write_development_report(report_path, result.report)
    return 0


def _write_development_report(path: Path, report: Any) -> None:
    serializer = getattr(report, "to_json", None)
    if not callable(serializer):
        raise RuntimeError("local repair report is not serializable")
    encoded = json.dumps(
        serializer(),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
        allow_nan=False,
    ).encode("utf-8")
    destination = Path(path).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(
        f".{destination.name}.{secrets.token_hex(8)}.tmp"
    )
    descriptor = os.open(
        str(temporary), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600
    )
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def _sanitize_development_diagnostic(value: str) -> str:
    value = _PRIVATE_KEY.sub("[REDACTED_PRIVATE_KEY]", value)
    value = _PORTABLE_FILENAME.sub("[REDACTED_PORTABLE_FILENAME]", value)
    value = _UUID.sub("[REDACTED_JOB_ID]", value)
    return _TOKEN.sub("[REDACTED_TOKEN]", value)


def _write_development_failure_diagnostic(path: Path, error: BaseException) -> None:
    payload = {
        "status": "failed",
        "error": {
            "type": type(error).__name__,
            "message": _sanitize_development_diagnostic(str(error)),
            "traceback": _sanitize_development_diagnostic(
                "".join(traceback.format_exception(error))
            ),
        },
    }
    destination = Path(path).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f"{destination.name}.tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temporary.replace(destination)


def _argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="LektaRepairDev")
    parser.add_argument("--development-e2e", action="store_true", required=True)
    parser.add_argument("--claim-endpoint", required=True)
    parser.add_argument("--status-endpoint", required=True)
    parser.add_argument("--output-directory", type=Path, required=True)
    parser.add_argument("--state-directory", type=Path, required=True)
    parser.add_argument("--diagnostic-path", type=Path)
    parser.add_argument("--report-path", type=Path)
    return parser


def main(
    argv: Sequence[str] | None = None,
    *,
    executable_path: Path | None = None,
    runner: OneShotRunner | Any | None = None,
) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if len(arguments) == 2 and arguments[0] == "--self-test":
        from word_replica.runner.portable_entry import run_portable_self_test

        return run_portable_self_test(arguments[1])
    if "--development-e2e" not in arguments:
        return 2
    diagnostic_path: Path | None = None
    try:
        parsed = _argument_parser().parse_args(arguments)
        diagnostic_path = parsed.diagnostic_path
        frozen_path = Path(
            executable_path
            if executable_path is not None
            else (sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
        )
        return run_development_e2e(
            frozen_path,
            claim_endpoint=parsed.claim_endpoint,
            status_endpoint=parsed.status_endpoint,
            output_dir=parsed.output_directory,
            state_dir=parsed.state_directory,
            report_path=parsed.report_path,
            runner=runner,
        )
    except (SystemExit, Exception) as error:
        if diagnostic_path is not None:
            try:
                _write_development_failure_diagnostic(diagnostic_path, error)
            except OSError:
                pass
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
