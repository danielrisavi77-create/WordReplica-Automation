"""Small fail-closed HTTPS transport for the one-shot Lekta runner."""
from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener


MAX_JSON_RESPONSE_BYTES = 1024 * 1024


class TransportError(RuntimeError):
    """A network response was unavailable, malformed or outside its limits."""


class _HttpsOnlyRedirectHandler(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        _require_https(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _default_open_request(request: Request, *, timeout: float):
    return build_opener(_HttpsOnlyRedirectHandler()).open(request, timeout=timeout)


def _require_https(url: str) -> None:
    parsed = urlsplit(url)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise TransportError("URL must use HTTPS")


def _read_bounded(response: Any, maximum_bytes: int) -> bytes:
    if isinstance(maximum_bytes, bool) or not isinstance(maximum_bytes, int) or maximum_bytes < 1:
        raise TransportError("invalid response size limit")
    declared = response.headers.get("Content-Length")
    if declared is not None:
        try:
            declared_size = int(declared)
        except (TypeError, ValueError) as exc:
            raise TransportError("invalid response length") from exc
        if declared_size < 0:
            raise TransportError("invalid response length")
        if declared_size > maximum_bytes:
            raise TransportError("response is too large")
    body = response.read(maximum_bytes + 1)
    if len(body) > maximum_bytes:
        raise TransportError("response is too large")
    return body


@dataclass(slots=True)
class HttpsTransport:
    open_request: Callable[..., Any] = _default_open_request
    timeout_seconds: float = 15.0

    def post_json(self, url: str, payload: dict) -> dict:
        _require_https(url)
        request = Request(
            url,
            data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
            headers={"Accept": "application/json", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with self.open_request(request, timeout=self.timeout_seconds) as response:
                body = _read_bounded(response, MAX_JSON_RESPONSE_BYTES)
        except TransportError:
            raise
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            raise TransportError("HTTPS request failed") from exc
        try:
            value = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise TransportError("invalid JSON response") from exc
        if not isinstance(value, dict):
            raise TransportError("JSON response must be an object")
        return value

    def get_bytes(self, url: str, *, maximum_bytes: int) -> bytes:
        _require_https(url)
        request = Request(url, headers={"Accept": "application/octet-stream"}, method="GET")
        try:
            with self.open_request(request, timeout=self.timeout_seconds) as response:
                return _read_bounded(response, maximum_bytes)
        except TransportError:
            raise
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            raise TransportError("HTTPS request failed") from exc
