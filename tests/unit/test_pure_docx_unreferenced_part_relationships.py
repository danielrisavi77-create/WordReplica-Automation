"""A part the source related but never referenced still needs its relationship.

Packages routinely carry a part that a relationship reaches and no element
points at: an image left behind by editing, or the drawing half of a SmartArt
graphic, which is reached from `word/_rels/document.xml.rels` while `dgm:relIds`
names only the data, layout, colours and quick-style parts.

The rebuild kept the part and dropped the relationship, so the part became an
orphan -- present in the package with nothing reaching it. That is a difference
from the source, which had the relationship; and an orphan is what G10 reports.
It was the largest remaining class of silent loss in the corpus: 7 of 19,
alongside 3 more documents whose image part disappeared entirely.

Restoring the relationship reproduces what the source actually was. The part was
not displayed there either -- the point is not to make a picture appear, it is
to stop changing the package while claiming to have preserved it.

The relationship id is allocated fresh, which is safe precisely because nothing
refers to it by id.
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
from word_replica.qa.preservation import g10_projection
from word_replica.services.rebuild import RebuildService

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753"
    "de0000000c4944415408d763f8cfc00000030101003c2f6e5c0000000049454e"
    "44ae426082"
)

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")
PLAIN_BODY = "<w:p><w:r><w:t>text</w:t></w:r></w:p>"


def _write(path, *, doc_rels: str, body: str = PLAIN_BODY, images=("word/media/image1.png",)):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{REL}">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">{doc_rels}</Relationships>'.encode(
                "utf-8"
            ),
        )
        archive.writestr("word/document.xml", document)
        for image in images:
            archive.writestr(image, PNG)
    return path


IMAGE_REL = f'<Relationship Id="rId9" Type="{REL}/image" Target="media/image1.png"/>'


@pytest.fixture
def stale_image(tmp_path):
    """An image with a relationship and nothing in the body pointing at it."""
    return _write(tmp_path / "stale.docx", doc_rels=IMAGE_REL)


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


def _rels(path, part="word/_rels/document.xml.rels"):
    with ZipFile(path) as archive:
        if part not in archive.namelist():
            return []
        return list(etree.fromstring(archive.read(part)))


def test_the_image_part_still_travels(tmp_path, stale_image):
    with ZipFile(_rebuild(tmp_path, stale_image)) as archive:
        assert any(name.startswith("word/media/") for name in archive.namelist())


def test_the_relationship_that_reached_it_comes_back(tmp_path, stale_image):
    nodes = _rels(_rebuild(tmp_path, stale_image, "-rel"))

    images = [node for node in nodes if (node.get("Type") or "").endswith("/image")]
    assert [node.get("Target") for node in images] == ["media/image1.png"]


def test_the_part_is_no_longer_an_orphan(tmp_path, stale_image):
    assert g10_projection(_rebuild(tmp_path, stale_image, "-orphan"))["orphan_parts"] == []


def test_g10_sees_the_package_as_preserved(tmp_path, stale_image):
    before = g10_projection(stale_image)
    after = g10_projection(_rebuild(tmp_path, stale_image, "-g10"))

    assert after["orphan_parts"] == before["orphan_parts"]
    assert after["part_kinds"].get("image/png") == before["part_kinds"].get("image/png")
    assert after["dangling_relationships"] == []


# --- the cases this must not disturb -------------------------------------------

def test_a_referenced_image_does_not_gain_a_second_relationship(tmp_path):
    body = (
        "<w:p><w:r><w:drawing><wp:inline "
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )
    source = _write(tmp_path / "used.docx", doc_rels=IMAGE_REL, body=body)

    nodes = _rels(_rebuild(tmp_path, source, "-used"))

    images = [node for node in nodes if (node.get("Type") or "").endswith("/image")]
    assert len(images) == 1, "the reference already created one; a second would be litter"


def test_a_document_with_no_stray_parts_gains_no_relationships(tmp_path):
    source = _write(tmp_path / "clean.docx", doc_rels="", images=())

    nodes = _rels(_rebuild(tmp_path, source, "-clean"))

    assert [node for node in nodes if (node.get("Type") or "").endswith("/image")] == []


def test_nothing_points_at_a_part_that_is_not_there(tmp_path):
    # The source relates an image whose bytes are missing. Restoring that
    # relationship would trade an orphan for a dangling one, and Word repairs
    # the second.
    source = _write(tmp_path / "missing.docx", doc_rels=IMAGE_REL, images=())

    output = _rebuild(tmp_path, source, "-missing")

    assert g10_projection(output)["dangling_relationships"] == []
