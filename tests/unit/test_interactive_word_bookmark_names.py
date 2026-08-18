from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
import pytest

from word_replica.renderers.interactive_word import (
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
