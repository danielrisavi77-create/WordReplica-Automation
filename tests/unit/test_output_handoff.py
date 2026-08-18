import os

from word_replica.services.output_handoff import open_output_for_user


def test_open_output_for_user_uses_the_default_windows_association(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(os, "startfile", lambda path: calls.append(path), raising=False)
    target = tmp_path / "Kalogjera - seminar Havel-popravljeno.docx"
    target.write_bytes(b"verified output")

    open_output_for_user(target)

    assert calls == [str(target)]


def test_open_output_for_user_never_imports_word_com():
    # A regression guard: this module must stay free of any win32com/pywin32
    # import, since it must never obtain or close a Word instance itself.
    import inspect

    import word_replica.services.output_handoff as module

    source = inspect.getsource(module)
    assert "win32com" not in source
    assert "pythoncom" not in source
