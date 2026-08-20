"""Pictures must come back into the document body, not just the package.

Verified directly before writing this: rebuilding
`05_images_inline_floating.docx` produced a package still containing
`word/media/image1.png` and a `word/document.xml` with no `w:drawing` at all.
Every rebuilt document lost every picture from its visible content while the
image bytes sat in the package as litter. It was the single largest remaining
G10 loss -- `wordprocessingDrawing` absent from 13 of 34 sampled documents --
and the renderer said so honestly through
`PURE_DOCX_ASSET_POSITION_UNAVAILABLE`.

The approach is to preserve the original `w:drawing` verbatim and remap its
relationship ids, rather than rebuild DrawingML by hand. Verbatim keeps crop,
rotation, wrapping, effects and anchor geometry exactly as authored; hand-built
markup would keep only what the model happens to represent.

Remapping is what makes it safe. The inline-fragment mechanism refuses any
fragment carrying an `r:` reference, because a renumbered id points at the
wrong part or none at all. Here the reference *can* be resolved -- the parser
knows which part each id addressed and the renderer owns the new package -- so
the id is rewritten rather than the fragment refused. An id that cannot be
resolved is still refused.
"""
from zipfile import ZipFile

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
from word_replica.qa.preservation import build_preservation_gate, g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


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


def _document(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml"))


def _media(path):
    with ZipFile(path) as archive:
        return sorted(n for n in archive.namelist() if n.startswith("word/media/"))


@pytest.fixture
def rebuilt(tmp_path, corpus_dir):
    source = corpus_dir / "05_images_inline_floating.docx"
    return source, _rebuild(tmp_path, source)


# --- the picture is in the document, not only in the package ------------------

def test_the_drawing_element_is_emitted(rebuilt):
    _source, output = rebuilt

    assert _document(output).findall(f".//{{{W}}}drawing"), "no w:drawing in the rebuilt body"


def test_the_media_part_still_travels_with_it(rebuilt):
    source, output = rebuilt

    assert _media(output) == _media(source)


def test_the_drawing_references_a_relationship_that_exists(rebuilt):
    # The whole risk of this change: a reference pointing at nothing produces a
    # document Word repairs on open, which is worse than a missing picture.
    _source, output = rebuilt

    assert g10_projection(output)["dangling_relationships"] == []


def test_the_image_is_no_longer_an_orphan_part(rebuilt):
    _source, output = rebuilt

    assert "image/png" not in g10_projection(output)["orphan_parts"]


def test_the_relationship_id_is_remapped_to_the_new_package(rebuilt):
    _source, output = rebuilt
    document = _document(output)

    embeds = {
        node.get(f"{{{R}}}embed")
        for node in document.iter()
        if isinstance(node.tag, str) and node.get(f"{{{R}}}embed")
    }
    assert embeds, "no r:embed survived"

    with ZipFile(output) as archive:
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
    declared = {node.get("Id") for node in rels}
    assert embeds <= declared, f"{embeds - declared} referenced but not declared"


# --- what verbatim preservation buys ------------------------------------------

def test_anchor_geometry_survives_verbatim(tmp_path, corpus_dir):
    # Hand-built DrawingML would keep only what the model represents. Verbatim
    # keeps the authored geometry, so an anchored picture stays anchored.
    source = corpus_dir / "05_images_inline_floating.docx"
    before, after = _document(source), _document(_rebuild(tmp_path, source, "-geom"))

    def _kinds(root):
        return sorted(
            node.tag.rsplit("}", 1)[-1]
            for node in root.iter()
            if isinstance(node.tag, str) and node.tag.rsplit("}", 1)[-1] in {"inline", "anchor", "extent"}
        )

    assert _kinds(after) == _kinds(before)


def test_g10_no_longer_reports_the_drawing_namespaces_as_lost(tmp_path, corpus_dir):
    source = corpus_dir / "05_images_inline_floating.docx"
    before = g10_projection(source)
    after = g10_projection(_rebuild(tmp_path, source, "-ns"))

    for uri in before["namespaces"]:
        if "wordprocessingDrawing" in uri or uri.endswith("/drawingml/2006/picture"):
            assert uri in after["namespaces"], f"{uri} lost"


def test_the_rebuild_is_still_parseable(tmp_path, corpus_dir):
    output = _rebuild(tmp_path, corpus_dir / "05_images_inline_floating.docx", "-parse")

    assert DocxParser().parse(output).drawings


def test_a_second_round_trip_is_stable(tmp_path, corpus_dir):
    # Each rebuild renumbers relationships. If remapping were not idempotent the
    # document would degrade a little on every pass.
    once = _rebuild(tmp_path, corpus_dir / "05_images_inline_floating.docx", "-1")
    twice = _rebuild(tmp_path, once, "-2")

    assert build_preservation_gate(once, twice).passed is True


def test_a_document_with_no_pictures_gains_none(tmp_path, corpus_dir):
    output = _rebuild(tmp_path, corpus_dir / "01_plain_text.docx", "-plain")

    assert not _document(output).findall(f".//{{{W}}}drawing")
    assert not _media(output)


# --- pictures that live outside the main document part ------------------------

def test_a_picture_inside_a_comment_is_not_orphaned(tmp_path, corpus_dir):
    """A relationship id means different things in different parts.

    The capture resolved every r: reference against word/document.xml.rels, so
    a picture inside a comment or a header -- whose id lives in that part's own
    .rels -- could not be resolved and the fragment was refused. The media part
    still travelled, leaving an image in the package that nothing displayed.
    """
    from zipfile import ZIP_DEFLATED
    from word_replica.qa.preservation import g10_projection

    base = corpus_dir / "10_comments_tracked_changes.docx"
    with ZipFile(base) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}

    comments = members["word/comments.xml"].decode("utf-8")
    drawing = (
        '<w:r><w:drawing>'
        '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<wp:extent cx="914400" cy="914400"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:blipFill><a:blip '
        'r:embed="rIdPic1" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        "</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )
    assert "</w:p></w:comment>" in comments, comments[:200]
    members["word/comments.xml"] = comments.replace(
        "</w:p></w:comment>", f"{drawing}</w:p></w:comment>", 1
    ).encode("utf-8")

    members["word/media/commentimage.png"] = b"\x89PNG\r\n\x1a\n" + b"c" * 40
    members["word/_rels/comments.xml.rels"] = (
        b'<?xml version="1.0"?><Relationships '
        b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rIdPic1" Target="media/commentimage.png" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>'
        b"</Relationships>"
    )
    members["[Content_Types].xml"] = members["[Content_Types].xml"].replace(
        b"</Types>", b'<Default Extension="png" ContentType="image/png"/></Types>'
    )
    enriched = tmp_path / "commentpic.docx"
    with ZipFile(enriched, "w", ZIP_DEFLATED) as out:
        for name, data in members.items():
            out.writestr(name, data)

    assert g10_projection(enriched)["orphan_parts"] == [], "the fixture itself must be sound"

    projection = g10_projection(_rebuild(tmp_path, enriched, "-commentpic"))

    assert "image/png" not in projection["orphan_parts"]
    assert projection["dangling_relationships"] == []


def test_the_same_relationship_id_resolves_per_part(tmp_path):
    """A relationship id is only meaningful relative to the part that declares it.

    Measured on a real corpus document before this was built: rId1 addressed
    word/styles.xml from word/document.xml.rels and word/media/image1.png from
    word/_rels/comments.xml.rels. Resolving without knowing whose id it is can
    only guess, and a guess rewrites a reference to the wrong part.
    """
    from zipfile import ZIP_DEFLATED
    from word_replica.opc.package_reader import DocxPackage
    from word_replica.parser.parser import _reference_targets

    def _rels(target: str, kind: str) -> bytes:
        return (
            '<?xml version="1.0"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Target="{target}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"/>'
            "</Relationships>"
        ).encode("utf-8")

    path = tmp_path / "collide.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", b"<Types/>")
        archive.writestr("_rels/.rels", b"<Relationships/>")
        archive.writestr("word/document.xml", b"<document/>")
        archive.writestr("word/comments.xml", b"<comments/>")
        archive.writestr("word/styles.xml", b"<styles/>")
        archive.writestr("word/media/image1.png", b"png")
        archive.writestr("word/_rels/document.xml.rels", _rels("styles.xml", "styles"))
        archive.writestr("word/_rels/comments.xml.rels", _rels("media/image1.png", "image"))

    with DocxPackage.open(path) as package:
        from_document = _reference_targets(package, "word/document.xml")["rId1"]
        from_comments = _reference_targets(package, "word/comments.xml")["rId1"]

    # (part, relationship type) -- the type travels because an OLE object
    # reached through an image relationship is not the same document.
    assert from_document[0] == "word/styles.xml"
    assert from_document[1].endswith("/styles")
    assert from_comments[0] == "word/media/image1.png"
    assert from_comments[1].endswith("/image")
