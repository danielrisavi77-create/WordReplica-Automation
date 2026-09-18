"""A zip directory entry is not an OPC part.

Zips may carry zero-length entries whose names end in "/" to record a folder.
Many writers emit them -- LibreOffice and Java-based tooling routinely do --
and Word reads such packages without complaint. `DocxPackage.parts` is built
straight from `namelist()`, so those entries were being handed back by
`iter_parts` as though they were parts.

The damage that does is out of all proportion to the cause. `theme_parts` is
collected with `iter_parts("word/theme/")` and the font scheme is read from the
first entry:

    data = next(iter(theme_parts.values()), None)
    if not data:
        return {}

`word/theme/` sorts before `word/theme/theme1.xml`, so the first entry is the
empty directory, `data` is empty, and the function gives up before looking at
the theme that is right there. Every theme font in the document then fails to
resolve.

That is the whole of the G2 class in the corpus: 11 of 260 documents, every one
of them "output invented a value", `runs/font_ascii` expected <missing> and got
Calibri. Nothing was invented. The rebuild writes no directory entries, so the
rebuilt package resolves its theme fonts correctly and the *source* is the one
that was read wrong -- the gate was reporting the parser's blind spot as a
renderer defect.

`iter_parts("word/media/")` had the same entry handed to it, and would have
tried to extract a folder as an image.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.opc.package_reader import DocxPackage
from word_replica.parser.parser import DocxParser

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-'
    'officedocument.theme+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")
DOC_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="theme/theme1.xml" Type="{REL}/theme"/>'
    "</Relationships>"
).encode("utf-8")
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
THEME = (
    f'<?xml version="1.0"?><a:theme xmlns:a="{A}" name="Office">'
    "<a:themeElements><a:fontScheme name=\"Office\">"
    '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Corbel"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    "</a:fontScheme></a:themeElements></a:theme>"
).encode("utf-8")


def _write(path, *, with_directory_entries: bool):
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        if with_directory_entries:
            # Exactly what LibreOffice and Java writers leave behind.
            archive.writestr("word/", b"")
            archive.writestr("word/theme/", b"")
            archive.writestr("word/media/", b"")
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/theme/theme1.xml", THEME)
    return path


@pytest.fixture
def with_folders(tmp_path):
    return _write(tmp_path / "folders.docx", with_directory_entries=True)


@pytest.fixture
def without_folders(tmp_path):
    return _write(tmp_path / "flat.docx", with_directory_entries=False)


def test_a_directory_entry_is_not_returned_as_a_part(with_folders):
    with DocxPackage.open(with_folders) as package:
        assert package.iter_parts("word/theme/") == ["word/theme/theme1.xml"]


def test_media_folders_are_not_offered_as_assets_either(with_folders):
    with DocxPackage.open(with_folders) as package:
        assert package.iter_parts("word/media/") == []


def test_the_theme_fonts_resolve_despite_the_folder_entry(with_folders):
    scheme = DocxParser().parse(with_folders).extras["theme_font_scheme"]

    assert scheme.get("minorAscii") == "Corbel"
    assert scheme.get("majorAscii") == "Cambria"


def test_the_folder_entries_make_no_difference_at_all(with_folders, without_folders):
    parser = DocxParser()

    with_entries = parser.parse(with_folders).extras["theme_font_scheme"]
    without_entries = parser.parse(without_folders).extras["theme_font_scheme"]

    assert with_entries == without_entries


def test_a_package_with_no_theme_still_reports_none(tmp_path):
    path = tmp_path / "themeless.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/", b"")
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}"/>'.encode("utf-8"),
        )
        archive.writestr("word/document.xml", DOCUMENT)

    assert DocxParser().parse(path).extras["theme_font_scheme"] == {}


def test_a_theme_carrying_no_font_scheme_does_not_hide_one_that_does(tmp_path):
    # Picking the first theme part blindly is the same mistake in a different
    # coat: a themeOverride sorts first and may declare no fonts at all.
    path = tmp_path / "override.docx"
    override = (
        f'<?xml version="1.0"?><a:theme xmlns:a="{A}" name="Override">'
        "<a:themeElements/></a:theme>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/theme/aOverride.xml", override)
        archive.writestr("word/theme/theme1.xml", THEME)

    assert DocxParser().parse(path).extras["theme_font_scheme"].get("minorAscii") == "Corbel"
