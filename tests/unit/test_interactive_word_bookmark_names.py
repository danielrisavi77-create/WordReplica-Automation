from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
import pytest

import word_replica.renderers.interactive_word as interactive_word_module
from word_replica.renderers.interactive_word import (
    InteractiveWordController,
    _replace_with_retry,
    _restore_bookmark_names,
    sanitize_word_bookmark_name,
)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_valid_name_passes_through_unchanged():
    taken = set()
    assert sanitize_word_bookmark_name("Chapter1", taken=taken) == "Chapter1"
    assert taken == {"Chapter1"}


def test_real_world_slug_names_from_kasalo_thesis_are_sanitized():
    # Real bookmark names pulled from a source document that broke Word's
    # COM Bookmarks.Add with "Bad bookmark name" - hyphens, periods, and a
    # leading digit are never valid, even though the OOXML file format
    # itself has no such restriction.
    names = [
        "metodologija-istraživanja",
        "prilog-1.-anketni-upitnik",
        "zaključak",
    ]
    taken = set()
    for name in names:
        sanitized = sanitize_word_bookmark_name(name, taken=taken)
        assert sanitized[0].isalpha() or sanitized[0] == "_"
        assert all(ch.isascii() and (ch.isalnum() or ch == "_") for ch in sanitized)
        assert len(sanitized) <= 40


def test_sanitized_names_never_collide():
    taken = set()
    first = sanitize_word_bookmark_name("a-b", taken=taken)
    second = sanitize_word_bookmark_name("a.b", taken=taken)
    assert first != second


def test_sanitized_name_starts_with_letter_or_underscore_even_for_leading_digit():
    taken = set()
    sanitized = sanitize_word_bookmark_name("1st-result", taken=taken)
    assert sanitized[0].isalpha() or sanitized[0] == "_"


def test_restore_bookmark_names_rewrites_saved_docx_back_to_original(tmp_path):
    docx_path = tmp_path / "sample.docx"
    document_xml = (
        b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        b'<w:document xmlns:w="' + _W_NS.encode() + b'">'
        b"<w:body>"
        b'<w:p><w:bookmarkStart w:id="0" w:name="metodologija_istra_ivanja"/>'
        b"<w:r><w:t>Text</w:t></w:r>"
        b'<w:bookmarkEnd w:id="0"/></w:p>'
        b"</w:body></w:document>"
    )
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml)

    _restore_bookmark_names(docx_path, {"metodologija_istra_ivanja": "metodologija-istraživanja"})

    with ZipFile(docx_path, "r") as archive:
        patched = archive.read("word/document.xml")
    root = etree.fromstring(patched)
    names = [el.get(f"{{{_W_NS}}}name") for el in root.iter(f"{{{_W_NS}}}bookmarkStart")]
    assert names == ["metodologija-istraživanja"]


def test_restore_bookmark_names_is_a_noop_with_no_rewrites(tmp_path):
    docx_path = tmp_path / "sample.docx"
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", b"<root/>")
    before = docx_path.read_bytes()

    _restore_bookmark_names(docx_path, {})

    assert docx_path.read_bytes() == before


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


def test_resumed_save_uses_rotating_save_as_and_mirrors_stable_checkpoint(tmp_path):
    output = tmp_path / "resumed.docx"
    output.write_bytes(b"checkpoint")

    class FakeDocument:
        def __init__(self):
            self.save_calls = 0
            self.save_as_calls = []

        @property
        def FullName(self):
            raise AssertionError("save must not query Word for an already-known path")

        def Save(self):
            raise AssertionError("an open Word checkpoint must not be saved in place")

        def SaveAs2(self, path, FileFormat=None):
            saved = Path(path)
            self.save_as_calls.append((saved, FileFormat))
            saved.write_bytes(b"rotated checkpoint")

    document = FakeDocument()
    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = document
    controller._last_saved_path = output

    saved_path = controller.save(output)

    assert saved_path == output.resolve()
    assert document.save_calls == 0
    assert len(document.save_as_calls) == 1
    active_path, file_format = document.save_as_calls[0]
    assert active_path != output
    assert file_format == 16
    assert output.read_bytes() == b"rotated checkpoint"
    assert controller._last_saved_path == active_path
    assert controller._checkpoint_delivery_path == output.resolve()


def test_resumed_checkpoint_rotates_to_a_new_word_path_instead_of_saving_in_place(tmp_path):
    output = tmp_path / "resumed.docx"
    output.write_bytes(b"checkpoint")

    class FakeDocument:
        def __init__(self):
            self.save_as_calls = []

        def Save(self):
            raise AssertionError("an open Word checkpoint must not be saved in place")

        def SaveAs2(self, path, FileFormat=None):
            saved = Path(path)
            self.save_as_calls.append((saved, FileFormat))
            saved.write_bytes(b"rotated checkpoint")

    document = FakeDocument()
    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = document
    controller._last_saved_path = output

    saved_path = controller.save(output)

    assert saved_path == output.resolve()
    active_path, file_format = document.save_as_calls[0]
    assert active_path != output
    assert active_path.parent == output.parent
    assert active_path.suffix == ".docx"
    assert output.read_bytes() == b"rotated checkpoint"
    assert file_format == 16
    assert controller._last_saved_path == active_path


def test_rotating_checkpoint_path_stays_below_word_255_character_limit():
    destination = Path("C:/") / ("very-long-folder/" * 20) / "reconstructed.docx"

    rotating = interactive_word_module._rotating_checkpoint_path(destination)

    assert rotating.parent != destination.parent
    assert len(str(rotating)) < 255
    assert rotating.name.startswith("wr-")
    assert rotating.suffix == ".docx"


def test_rotating_save_survives_transient_lock_while_cleaning_previous_shadow(tmp_path, monkeypatch):
    output = tmp_path / "resumed.docx"
    output.write_bytes(b"old checkpoint")
    previous_shadow = tmp_path / "wr-0123456789abcdef.docx"
    previous_shadow.write_bytes(b"previous shadow")

    class FakeDocument:
        def SaveAs2(self, path, FileFormat=None):
            Path(path).write_bytes(b"new checkpoint")

    real_unlink = Path.unlink

    def locked_unlink(self, *args, **kwargs):
        if self == previous_shadow:
            raise PermissionError(32, "file is being used by another process")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", locked_unlink)
    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = FakeDocument()
    controller._last_saved_path = previous_shadow
    controller._checkpoint_delivery_path = output

    saved_path = controller.save(output)

    assert saved_path == output.resolve()
    assert output.read_bytes() == b"new checkpoint"
    assert previous_shadow.exists()


def test_visible_controller_does_not_toggle_or_reactivate_word_during_rotating_save(tmp_path):
    output = tmp_path / "resumed.docx"
    output.write_bytes(b"checkpoint")
    calls = []

    class FakeDocument:
        def SaveAs2(self, path, FileFormat=None):
            calls.append(("document.SaveAs2", FileFormat))
            Path(path).write_bytes(b"saved")

        def Activate(self):
            calls.append("document.Activate")

    class FakeApplication:
        def __init__(self):
            object.__setattr__(self, "Visible", True)

        def __setattr__(self, name, value):
            if name == "Visible":
                calls.append(("Visible", value))
            object.__setattr__(self, name, value)

        def Activate(self):
            calls.append("application.Activate")

    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = FakeDocument()
    controller.application = FakeApplication()
    controller._last_saved_path = output

    controller.save(output)

    assert calls == [("document.SaveAs2", 16)]


def test_visible_rotating_save_falls_back_to_hidden_word_after_rejected_calls(tmp_path, monkeypatch):
    output = tmp_path / "resumed.docx"
    output.write_bytes(b"checkpoint")
    calls = []

    class RejectedCall(RuntimeError):
        hresult = -2147418111

    class FakeApplication:
        def __init__(self):
            object.__setattr__(self, "Visible", True)

        def __setattr__(self, name, value):
            if name == "Visible":
                calls.append(("Visible", value))
            object.__setattr__(self, name, value)

    application = FakeApplication()

    class FakeDocument:
        def __init__(self):
            self.attempts = 0

        def SaveAs2(self, path, FileFormat=None):
            self.attempts += 1
            if application.Visible:
                raise RejectedCall("Call was rejected by callee.")
            Path(path).write_bytes(b"saved while hidden")

    monkeypatch.setattr(interactive_word_module.time, "sleep", lambda _seconds: None)
    document = FakeDocument()
    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = document
    controller.application = application
    controller._last_saved_path = output

    saved_path = controller.save(output)

    assert saved_path == output.resolve()
    assert output.read_bytes() == b"saved while hidden"
    assert document.attempts == 11
    assert calls == [("Visible", False), ("Visible", True)]


def test_bookmark_names_are_restored_only_after_close_not_during_save(tmp_path, monkeypatch):
    # Regression: SaveAs2 leaves the destination file locked by Word for as
    # long as the document stays open (confirmed live: 10/10 rename attempts
    # over an in-process Word session all failed with WinError 5, and only
    # succeeded once the document was closed). Restoring bookmark names via a
    # rename-based rewrite must therefore happen after close(), never during
    # save() while the document is still open.
    restore_calls = []
    monkeypatch.setattr(
        interactive_word_module,
        "_restore_bookmark_names",
        lambda path, rewrites: restore_calls.append((path, dict(rewrites))),
    )

    class FakeDocument:
        def SaveAs2(self, path, FileFormat=None):
            Path(path).write_bytes(b"saved")

        def Close(self, save_changes):
            pass

    class FakeApplication:
        def Quit(self):
            pass

    controller = InteractiveWordController.for_testing(active_range=None)
    controller.document = FakeDocument()
    controller.application = FakeApplication()
    controller._bookmark_name_rewrites = {"com_safe": "original-name"}

    output = tmp_path / "output.docx"
    controller.save(output)
    assert restore_calls == []

    controller.close()
    assert restore_calls == [(output, {"com_safe": "original-name"})]


def test_restore_bookmark_names_cleans_up_temp_file_when_rename_never_succeeds(tmp_path, monkeypatch):
    docx_path = tmp_path / "sample.docx"
    document_xml = (
        b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        b'<w:document xmlns:w="' + _W_NS.encode() + b'">'
        b"<w:body>"
        b'<w:p><w:bookmarkStart w:id="0" w:name="com_safe"/>'
        b"<w:r><w:t>Text</w:t></w:r>"
        b'<w:bookmarkEnd w:id="0"/></w:p>'
        b"</w:body></w:document>"
    )
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", document_xml)

    def always_denied(self, target):
        raise PermissionError(5, "Access is denied")

    import word_replica.renderers.interactive_word as interactive_word_module

    monkeypatch.setattr(Path, "replace", always_denied)
    monkeypatch.setattr(interactive_word_module, "_replace_with_retry", lambda src, dst: src.replace(dst))

    with pytest.raises(PermissionError):
        _restore_bookmark_names(docx_path, {"com_safe": "original-name"})

    assert not docx_path.with_suffix(".docx.tmp").exists()
