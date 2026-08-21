"""A part kept byte-for-byte keeps its relationships too.

styles.xml, numbering.xml and settings.xml are not rebuilt from the model -- the
renderer writes the source bytes back unchanged. Their `.rels` files were not
written back, and that is worse than losing content: the preserved bytes still
carry the original `r:id`, so the output ships a reference to a relationship
that no longer exists.

Found in the corpus as `test_extra_image.docx`, a picture bullet. Its
numbering.xml holds `<v:imagedata r:id="rId1"/>` and its
`word/_rels/numbering.xml.rels` maps rId1 to `media/image1.jpeg`. The rebuild
kept numbering.xml and the image part, dropped the .rels, and produced a
document with a dangling reference and an orphaned image -- which is a document
Word repairs on open, not merely one that lost a bullet.

settings.xml already had a narrow version of this fix that restored only its
*external* relationships; that special case is the same defect noticed once and
solved for one part.

The .rels can be restored verbatim precisely because the part it belongs to is
verbatim: the ids in it still match the ids in the bytes. What must be checked
is the other end -- an internal target whose part did not make it into the
package would turn one dangling reference into another, so it is reported
instead.
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
VML = "urn:schemas-microsoft-com:vml"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="jpeg" ContentType="image/jpeg"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.numbering+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")
DOC_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="numbering.xml" Type="{REL}/numbering"/>'
    "</Relationships>"
).encode("utf-8")

# A picture bullet: the reference lives in numbering.xml, not in the body.
NUMBERING = (
    f'<?xml version="1.0"?><w:numbering xmlns:w="{W}" xmlns:r="{REL}" xmlns:v="{VML}" '
    'xmlns:o="urn:schemas-microsoft-com:office:office">'
    '<w:numPicBullet w:numPicBulletId="0"><w:pict>'
    '<v:shape id="_x0000_i1025" style="width:9pt;height:9pt">'
    '<v:imagedata r:id="rId1" o:title="bullet"/>'
    "</v:shape></w:pict></w:numPicBullet>"
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
    '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>'
    "</w:lvl></w:abstractNum>"
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    "</w:numbering>"
).encode("utf-8")
NUMBERING_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Type="{REL}/image" Target="media/image1.jpeg"/>'
    "</Relationships>"
).encode("utf-8")

# A one-pixel JPEG: real bytes, so the image part is a genuine part.
JPEG = bytes.fromhex(
    "ffd8ffe000104a46494600010100000100010000ffdb004300ff"
    "ffffffffffffffffffffffffffffffffffffffffffffffffffff"
    "ffffffffffffffffffffffffffffffffffffffffffffffffffff"
    "ffffffffffffffffffffffffffffffffffffffffffffffffff"
    "ffc00011080001000103012200021101031101ffc4001f00000105"
    "01010101010100000000000000000102030405060708090a0bffc4"
    "00b5100002010303020403050504040000017d01020300041105122"
    "1314106136151607122714328191a1082342b1c11552d1f02433627"
    "2820ffda000c03010002110311003f00fbfeffd9"
)

DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
    "<w:r><w:t>bulleted</w:t></w:r></w:p>"
    "<w:sectPr/></w:body></w:document>"
).encode("utf-8")


@pytest.fixture
def source(tmp_path):
    path = tmp_path / "picture-bullet.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/numbering.xml", NUMBERING)
        archive.writestr("word/_rels/numbering.xml.rels", NUMBERING_RELS)
        archive.writestr("word/media/image1.jpeg", JPEG)
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


def _names(path):
    with ZipFile(path) as archive:
        return set(archive.namelist())


def test_the_numbering_relationships_part_survives(tmp_path, source):
    assert "word/_rels/numbering.xml.rels" in _names(_rebuild(tmp_path, source))


def test_the_picture_bullet_still_resolves(tmp_path, source):
    output = _rebuild(tmp_path, source, "-resolve")
    with ZipFile(output) as archive:
        numbering = etree.fromstring(archive.read("word/numbering.xml"))
        rels = etree.fromstring(archive.read("word/_rels/numbering.xml.rels"))

    used = {
        node.get(f"{{{REL}}}id")
        for node in numbering.iter()
        if node.get(f"{{{REL}}}id")
    }
    declared = {node.get("Id") for node in rels}
    assert used, "the preserved numbering.xml lost its picture bullet reference"
    assert used <= declared, f"dangling reference: {used - declared}"


def test_the_bullet_image_is_not_orphaned(tmp_path, source):
    # G10 reports a part nothing points at, which is exactly what dropping the
    # .rels leaves behind.
    assert g10_projection(_rebuild(tmp_path, source, "-orphan"))["orphan_parts"] == []


def test_no_reference_is_left_dangling(tmp_path, source):
    assert g10_projection(_rebuild(tmp_path, source, "-dangle"))["dangling_relationships"] == []


def test_g10_sees_the_package_as_preserved(tmp_path, source):
    output = _rebuild(tmp_path, source, "-g10")
    before, after = g10_projection(source), g10_projection(output)

    assert after["orphan_parts"] == before["orphan_parts"]
    assert after["part_kinds"].get("image/jpeg") == before["part_kinds"].get("image/jpeg")


def test_a_document_whose_numbering_has_no_relationships_gains_no_rels_part(tmp_path):
    path = tmp_path / "plain-numbering.docx"
    numbering = (
        f'<?xml version="1.0"?><w:numbering xmlns:w="{W}">'
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
        '<w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        "</w:numbering>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/numbering.xml", numbering)

    assert "word/_rels/numbering.xml.rels" not in _names(_rebuild(tmp_path, path, "-plain"))
