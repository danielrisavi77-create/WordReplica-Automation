import pytest

from word_replica.domain.enums import RendererChoice
from word_replica.domain.errors import RendererUnavailableError
from word_replica.renderers.word_com import choose_renderer_name, word_available


def test_auto_falls_back_to_docx_when_word_unavailable():
    assert choose_renderer_name(RendererChoice.AUTO, word_is_available=False) == "docx"


def test_explicit_word_requires_word():
    with pytest.raises(RendererUnavailableError):
        choose_renderer_name(RendererChoice.WORD, word_is_available=False)


def test_non_windows_reports_word_unavailable(monkeypatch):
    monkeypatch.setattr("word_replica.renderers.word_com.sys.platform", "linux")
    assert word_available() is False
