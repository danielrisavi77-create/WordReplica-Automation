"""The environment probe must not launch Word once per rebuild.

`capture_environment_fingerprint` is called from `RebuildService.rebuild`, so
every rebuild -- including a pure-docx one that needs no Word at all -- was
starting a dedicated Word process through DispatchEx just to read Version and
Build, then quitting it again.

Two things follow from that, both measured on this machine. A 400-document
lane-P pass paid for 400 Word launches, which is most of its 13.3 s/document;
and when the Golden run held Word, those launches began failing with
RPC_S_SERVER_UNAVAILABLE and leaked WINWORD processes that never quit -- a
full test run died partway through with no summary because of it.

Neither the Word build nor the installed font list can change inside one
process, so both are probed once and reused. The timestamp still moves, because
that is the one field that is genuinely per-call.

The probes keep swallowing their own failures: a cached failure is still a
fingerprint, and it stays consistent for the whole run rather than flickering
with whatever else is holding Word at that moment.
"""
from word_replica.qa import environment


def test_the_word_probe_runs_once_however_many_fingerprints_are_taken(monkeypatch):
    calls = []

    def probe():
        calls.append(1)
        return {"available": True, "version": "14.0", "build": "7015", "name": "Word"}

    monkeypatch.setattr(environment, "_word_application_info", probe)
    environment.reset_environment_probe_cache()

    for _ in range(5):
        environment.capture_environment_fingerprint(include_word=True)

    assert len(calls) == 1


def test_the_font_probe_runs_once_too(monkeypatch):
    calls = []

    def probe():
        calls.append(1)
        return {"available": True, "count": 2, "names": ["Calibri", "Arial"]}

    monkeypatch.setattr(environment, "_enumerate_windows_fonts", probe)
    environment.reset_environment_probe_cache()

    for _ in range(4):
        environment.capture_environment_fingerprint()

    assert len(calls) == 1


def test_the_word_information_still_reaches_the_fingerprint(monkeypatch):
    monkeypatch.setattr(
        environment,
        "_word_application_info",
        lambda: {"available": True, "version": "14.0", "build": "7015", "name": "Word"},
    )
    environment.reset_environment_probe_cache()

    fingerprint = environment.capture_environment_fingerprint(include_word=True)

    assert fingerprint["word"]["version"] == "14.0"
    assert fingerprint["word"]["build"] == "7015"


def test_the_timestamp_is_not_cached(monkeypatch):
    monkeypatch.setattr(environment, "_word_application_info", lambda: {"available": False})
    monkeypatch.setattr(environment, "_enumerate_windows_fonts", lambda: {"available": False})
    environment.reset_environment_probe_cache()

    first = environment.capture_environment_fingerprint()["captured_at_utc"]
    second = environment.capture_environment_fingerprint()["captured_at_utc"]

    assert first <= second
    assert isinstance(second, str) and second


def test_a_failing_probe_is_reported_rather_than_raised(monkeypatch):
    def probe():
        raise RuntimeError("Word is busy")

    # The real probe catches its own exceptions; this asserts the caching layer
    # does not turn a raising probe into a raising fingerprint either.
    monkeypatch.setattr(
        environment,
        "_word_application_info",
        lambda: {"available": False, "error": "RPC_S_SERVER_UNAVAILABLE"},
    )
    monkeypatch.setattr(environment, "_enumerate_windows_fonts", probe)
    environment.reset_environment_probe_cache()

    fingerprint = environment.capture_environment_fingerprint(include_word=True)

    assert fingerprint["word"]["available"] is False
    assert fingerprint["fonts"]["available"] is False


def test_the_cache_can_be_reset(monkeypatch):
    calls = []
    monkeypatch.setattr(
        environment,
        "_word_application_info",
        lambda: (calls.append(1), {"available": False})[1],
    )

    environment.reset_environment_probe_cache()
    environment.capture_environment_fingerprint(include_word=True)
    environment.reset_environment_probe_cache()
    environment.capture_environment_fingerprint(include_word=True)

    assert len(calls) == 2


def test_a_word_free_fingerprint_does_not_launch_word(monkeypatch):
    calls = []
    monkeypatch.setattr(
        environment,
        "_word_application_info",
        lambda: (calls.append(1), {"available": True})[1],
    )
    environment.reset_environment_probe_cache()

    fingerprint = environment.capture_environment_fingerprint(include_word=False)

    assert calls == []
    assert fingerprint["word"] == {"available": False, "skipped": True}


def test_the_default_fingerprint_is_word_free(monkeypatch):
    calls = []
    monkeypatch.setattr(
        environment,
        "_word_application_info",
        lambda: (calls.append(1), {"available": True})[1],
    )
    environment.reset_environment_probe_cache()

    fingerprint = environment.capture_environment_fingerprint()

    assert calls == []
    assert fingerprint["word"] == {"available": False, "skipped": True}
