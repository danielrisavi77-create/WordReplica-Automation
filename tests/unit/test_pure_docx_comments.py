"""Comments must survive a rebuild.

Found by G10 on the corpus: `word/comments.xml` present in the source, absent
from the reconstruction. The parser reads comments into the model and the
renderer never wrote them back -- the only mention of comments.xml in
`pure_docx.py` was in an unrelated list of story parts.

So every comment in every document went through the pure-docx path and came out
the other side gone, and no gate could see it: G0 compares body text, which a
comment is not part of, and the comment part has no projection anywhere in
`qa/`.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from lxml import etree

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.parser.parser import DocxParser
from word_replica.qa.preservation import g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = (
    b'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.document.main+xml"/>'
    b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument'
    b'.wordprocessingml.comments+xml"/>'
    b"</Types>"
)
ROOT_RELS = (
    b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" Target="word/document.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    b"</Relationships>"
)
DOC_RELS = (
    b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId5" Target="comments.xml" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"/>'
    b"</Relationships>"
)
DOCUMENT = (
    '<?xml version="1.0"?><w:document '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p>"
    '<w:commentRangeStart w:id="1"/>'
    "<w:r><w:t>reviewed text</w:t></w:r>"
    '<w:commentRangeEnd w:id="1"/>'
    '<w:r><w:commentReference w:id="1"/></w:r>'
    "</w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")
COMMENTS = (
    '<?xml version="1.0"?><w:comments '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:comment w:id="1" w:author="Ana Kovac" w:date="2026-01-02T10:00:00Z" w:initials="AK">'
    "<w:p><w:r><w:t>Provjeri ovaj odlomak.</w:t></w:r></w:p>"
    "</w:comment></w:comments>"
).encode("utf-8")


@pytest.fixture
def source(tmp_path):
    path = tmp_path / "src.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/comments.xml", COMMENTS)
    return path


def _rebuild(tmp_path, source, suffix=""):
    result = RebuildService(app_root=tmp_path / f"app{suffix}").rebuild(
        source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )
    assert result.output_path is not None, result.reasons
    return result.output_path


def _part(path, name):
    with ZipFile(path) as archive:
        if name not in archive.namelist():
            return None
        return etree.fromstring(archive.read(name))


def test_the_comment_part_is_written(tmp_path, source):
    assert _part(_rebuild(tmp_path, source), "word/comments.xml") is not None


def test_the_comment_text_survives(tmp_path, source):
    comments = _part(_rebuild(tmp_path, source, "-text"), "word/comments.xml")

    assert "Provjeri ovaj odlomak." in "".join(comments.itertext())


def test_the_author_and_date_survive(tmp_path, source):
    comments = _part(_rebuild(tmp_path, source, "-meta"), "word/comments.xml")
    comment = comments.find(f"{{{W}}}comment")

    assert comment.get(f"{{{W}}}author") == "Ana Kovac"
    assert comment.get(f"{{{W}}}date") == "2026-01-02T10:00:00Z"
    assert comment.get(f"{{{W}}}id") == "1"


def test_the_body_still_points_at_the_comment(tmp_path, source):
    # A comment part with nothing referring to it is invisible in Word: the
    # text is in the file but no reader ever shows it.
    document = _part(_rebuild(tmp_path, source, "-ref"), "word/document.xml")

    references = document.findall(f".//{{{W}}}commentReference")
    assert [node.get(f"{{{W}}}id") for node in references] == ["1"]


def test_the_reference_stays_in_its_own_run(tmp_path, source):
    document = _part(_rebuild(tmp_path, source, "-run"), "word/document.xml")

    runs_with_reference = [
        run for run in document.iter(f"{{{W}}}r") if run.find(f"{{{W}}}commentReference") is not None
    ]
    assert len(runs_with_reference) == 1


def test_the_reviewed_text_is_untouched(tmp_path, source):
    assert "reviewed text" in DocxParser().parse(_rebuild(tmp_path, source, "-body")).plain_text()


def test_g10_no_longer_reports_the_comment_part_as_lost(tmp_path, source):
    before = g10_projection(source)
    after = g10_projection(_rebuild(tmp_path, source, "-g10"))

    kind = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    assert after["part_kinds"].get(kind) == before["part_kinds"].get(kind)
    assert "comments" in after["relationship_graph"]
    assert after["dangling_relationships"] == []


def test_a_document_without_comments_gains_no_comment_part(tmp_path):
    path = tmp_path / "bare.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES.replace(
            b'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-'
            b'officedocument.wordprocessingml.comments+xml"/>', b""
        ))
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", (
            '<?xml version="1.0"?><w:document '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        ).encode("utf-8"))

    assert _part(_rebuild(tmp_path, path, "-none"), "word/comments.xml") is None
