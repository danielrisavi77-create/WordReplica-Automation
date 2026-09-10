from word_replica.cli import (
    _default_one_shot_runner,
    _default_repair_package_service,
)
from word_replica.domain.enums import InteractiveSpeedMode, RendererChoice


def test_portable_lekta_runner_defaults_to_license_free_pure_docx(monkeypatch):
    monkeypatch.setattr("word_replica.runner.trust_store.load_trust_keys", lambda _path: {})

    runner = _default_one_shot_runner()

    assert runner.package_service.rebuild_options.renderer is RendererChoice.DOCX


def test_portable_pure_docx_service_is_silent_and_unthrottled(capsys):
    service = _default_repair_package_service("pure-docx", visible_preview=False)

    assert service.preview_interactive_options.speed_mode is InteractiveSpeedMode.MAXIMUM
    service.preview_sink("text", {"text": "private document text"})
    assert capsys.readouterr().out == ""
