from pathlib import Path

import pytest

from word_replica.services.interactive_rebuild import _replace_with_retry


def test_replace_with_retry_succeeds_immediately_when_no_contention(tmp_path):
    source = tmp_path / "a.tmp"
    destination = tmp_path / "a.docx"
    source.write_bytes(b"data")

    _replace_with_retry(source, destination)

    assert destination.read_bytes() == b"data"
    assert not source.exists()


def test_replace_with_retry_recovers_from_transient_permission_error(tmp_path, monkeypatch):
    source = tmp_path / "a.tmp"
    destination = tmp_path / "a.docx"
    source.write_bytes(b"data")
    calls = {"count": 0}
    real_replace = Path.replace

    def flaky_replace(self, target):
        calls["count"] += 1
        if calls["count"] < 3:
            raise PermissionError(5, "Access is denied")
        return real_replace(self, target)

    monkeypatch.setattr(Path, "replace", flaky_replace)

    _replace_with_retry(source, destination, attempts=5, delay_seconds=0)

    assert calls["count"] == 3
    assert destination.read_bytes() == b"data"


def test_replace_with_retry_raises_after_exhausting_attempts(tmp_path, monkeypatch):
    source = tmp_path / "a.tmp"
    destination = tmp_path / "a.docx"
    source.write_bytes(b"data")

    def always_denied(self, target):
        raise PermissionError(5, "Access is denied")

    monkeypatch.setattr(Path, "replace", always_denied)

    with pytest.raises(PermissionError):
        _replace_with_retry(source, destination, attempts=3, delay_seconds=0)
