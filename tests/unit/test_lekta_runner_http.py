from __future__ import annotations

import io
import json
from email.message import Message
from urllib.request import Request

import pytest

from word_replica.runner.http_transport import HttpsTransport, TransportError


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


def test_post_json_uses_https_json_and_a_bounded_response() -> None:
    captured: list[tuple[Request, float]] = []

    def open_request(request: Request, *, timeout: float) -> FakeResponse:
        captured.append((request, timeout))
        return FakeResponse(b'{"ok":true,"jobId":"job"}', content_length=25)

    transport = HttpsTransport(open_request=open_request, timeout_seconds=12.5)
    result = transport.post_json("https://lekta.example/claim", {"jobId": "job"})

    request, timeout = captured[0]
    assert request.full_url == "https://lekta.example/claim"
    assert request.get_method() == "POST"
    assert request.get_header("Content-type") == "application/json"
    assert json.loads(request.data or b"") == {"jobId": "job"}
    assert timeout == 12.5
    assert result == {"ok": True, "jobId": "job"}


@pytest.mark.parametrize("method", ["post", "get"])
def test_transport_rejects_non_https_before_network(method: str) -> None:
    called = False

    def open_request(_request: Request, *, timeout: float) -> FakeResponse:
        nonlocal called
        called = True
        return FakeResponse(b"{}")

    transport = HttpsTransport(open_request=open_request)
    with pytest.raises(TransportError):
        if method == "post":
            transport.post_json("http://lekta.example/claim", {})
        else:
            transport.get_bytes("file:///source.docx", maximum_bytes=10)
    assert called is False


def test_download_reads_at_most_one_byte_beyond_the_limit() -> None:
    body = b"123456"
    response = FakeResponse(body)
    transport = HttpsTransport(open_request=lambda *_a, **_k: response)

    with pytest.raises(TransportError, match="too large"):
        transport.get_bytes("https://storage.example/source", maximum_bytes=5)

    assert response._stream.tell() == 6


def test_declared_oversize_download_is_rejected_without_reading_body() -> None:
    response = FakeResponse(b"secret", content_length=6)
    transport = HttpsTransport(open_request=lambda *_a, **_k: response)

    with pytest.raises(TransportError, match="too large"):
        transport.get_bytes("https://storage.example/source", maximum_bytes=5)
