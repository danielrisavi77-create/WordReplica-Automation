from word_replica.qa import environment as environment_module
from word_replica.qa.environment import capture_environment_fingerprint


def test_capture_environment_fingerprint_has_expected_shape():
    fingerprint = capture_environment_fingerprint()
    assert set(fingerprint) == {"captured_at_utc", "os", "python", "locale", "word", "fonts"}
    assert fingerprint["os"]["system"]
    assert fingerprint["python"]["version"]
    assert "available" in fingerprint["word"]
    assert "available" in fingerprint["fonts"]


def test_word_application_info_reports_unavailable_off_windows(monkeypatch):
    monkeypatch.setattr(environment_module.sys, "platform", "linux")
    info = environment_module._word_application_info()
    assert info == {"available": False, "error": "not Windows"}


def test_enumerate_windows_fonts_reports_unavailable_off_windows(monkeypatch):
    monkeypatch.setattr(environment_module.sys, "platform", "linux")
    fonts = environment_module._enumerate_windows_fonts()
    assert fonts == {"available": False, "count": 0, "names": []}
