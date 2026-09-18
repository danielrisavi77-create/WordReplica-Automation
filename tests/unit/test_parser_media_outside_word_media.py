"""A picture is wherever its relationship says it is.

Media parts were collected by sweeping `word/media/`. Nothing in OOXML requires
a picture to live there -- the part name is whatever the relationship targets,
and writers do use other places. LibreOffice's fdo63685.docx keeps its picture
at `media/uidff77d259a14.jpeg`, at the root of the package.

A part outside the swept prefix never became an asset, so the drawing that
referenced it had nothing to resolve to, the part was never written, and the
picture was gone: G10 reported image/jpeg dropping from one to none along with
the image relationship.

Sweeping the folder stays -- it is what nearly every document does and it also
catches media that no drawing points at -- and parts reached by an image
relationship are collected as well, wherever they sit.
"""
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

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
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

JPEG = bytes.fromhex(
    "ffd8ffe000104a46494600010100000100010000ffdb004300"
    + "ff" * 64
    + "ffc00011080001000103012200021101031101"
    + "ffda000c03010002110311003f00bf"
    + "ffd9"
)


def _content_types(image_part: str) -> bytes:
    return (
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        f'<Override PartName="/{image_part}" ContentType="image/jpeg"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    ).encode("utf-8")


ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")

DRAWING = (
    "<w:p><w:r><w:drawing>"
    f'<wp:inline xmlns:wp="{WP}">'
    '<wp:extent cx="914400" cy="914400"/>'
    '<wp:docPr id="1" name="Picture 1"/>'
    f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}">'
    f'<pic:pic xmlns:pic="{PIC}">'
    '<pic:nvPicPr><pic:cNvPr id="0" name="p.jpeg"/><pic:cNvPicPr/></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
)


def _write(path, image_part: str, target: str):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{REL}">'
        f"<w:body>{DRAWING}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(image_part))
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            (
                f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
                f'<Relationship Id="rId7" Type="{REL}/image" Target="{target}"/>'
                "</Relationships>"
            ).encode("utf-8"),
        )
        archive.writestr("word/document.xml", document)
        archive.writestr(image_part, JPEG)
    return path


@pytest.fixture
def at_package_root(tmp_path):
    """The picture lives at media/, not word/media/ -- as fdo63685.docx does."""
    return _write(tmp_path / "root-media.docx", "media/picture.jpeg", "../media/picture.jpeg")


@pytest.fixture
def in_word_media(tmp_path):
    return _write(tmp_path / "word-media.docx", "word/media/picture.jpeg", "media/picture.jpeg")


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


def test_the_picture_reaches_the_model(at_package_root):
    parts = {asset.part_name for asset in DocxParser().parse(at_package_root).assets.values()}

    assert "media/picture.jpeg" in parts


def test_the_picture_bytes_travel(tmp_path, at_package_root):
    with ZipFile(_rebuild(tmp_path, at_package_root)) as archive:
        images = [n for n in archive.namelist() if n.lower().endswith(".jpeg")]

    assert images, "the picture was dropped"


def test_g10_still_counts_the_image(tmp_path, at_package_root):
    before = g10_projection(at_package_root)
    after = g10_projection(_rebuild(tmp_path, at_package_root, "-g10"))

    assert after["part_kinds"].get("image/jpeg") == before["part_kinds"].get("image/jpeg")


def test_the_image_relationship_survives(tmp_path, at_package_root):
    before = g10_projection(at_package_root)
    after = g10_projection(_rebuild(tmp_path, at_package_root, "-rel"))

    assert after["relationship_graph"].get("image") == before["relationship_graph"].get("image")


# --- the ordinary case is unchanged --------------------------------------------

def test_a_picture_in_word_media_is_unaffected(tmp_path, in_word_media):
    before = g10_projection(in_word_media)
    after = g10_projection(_rebuild(tmp_path, in_word_media, "-normal"))

    assert after["part_kinds"].get("image/jpeg") == before["part_kinds"].get("image/jpeg")


def test_it_is_collected_once_not_twice(in_word_media):
    model = DocxParser().parse(in_word_media)

    parts = [asset.part_name for asset in model.assets.values()]
    assert parts.count("word/media/picture.jpeg") == 1
