"""Two parts holding the same bytes are still two parts.

`BinaryAsset.asset_id` is derived from the content hash, and media parts were
collected with `model.assets.setdefault(asset.asset_id, asset)`. A package that
stores the same image twice under two names therefore kept only the first: the
second part never reached the model, so the rebuild never wrote it, and the
package came out one part short.

In OPC a part's identity is its name. Word keeps both, and it is not unusual for
a document to have them -- LibreOffice's fdo65833.docx carries
`word/media/image1.tmp` and `word/media/image2.png`, byte-identical at 6379
bytes, one referenced through `v:imagedata` and the other through `a:blip` in
the two halves of an mc:AlternateContent. The rebuild dropped image2.png and
G10 reported image/png going from two to one.

Ids stay content-derived, because DrawingRef.asset_id, the interactive
preflight and the asset files written to disk all key on them. Only a genuine
collision between two different parts is disambiguated, so every document
without one keeps exactly the ids it had.
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
    '<Override PartName="/word/media/image1.tmp" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")
DOC_RELS = (
    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
    f'<Relationship Id="rId4" Type="{REL}/image" Target="media/image1.tmp"/>'
    f'<Relationship Id="rId5" Type="{REL}/image" Target="media/image2.png"/>'
    "</Relationships>"
).encode("utf-8")
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:r="{REL}"><w:body>'
    "<w:p><w:r><w:t>two pictures, one picture's worth of bytes</w:t></w:r></w:p>"
    "<w:sectPr/></w:body></w:document>"
).encode("utf-8")


@pytest.fixture
def twins(tmp_path):
    path = tmp_path / "twins.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOC_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/media/image1.tmp", PNG)
        archive.writestr("word/media/image2.png", PNG)
    return path


@pytest.fixture
def single(tmp_path):
    path = tmp_path / "single.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
            f'<Relationship Id="rId4" Type="{REL}/image" Target="media/image1.tmp"/>'
            "</Relationships>".encode("utf-8"),
        )
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/media/image1.tmp", PNG)
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


def test_both_parts_reach_the_model(twins):
    model = DocxParser().parse(twins)

    assert sorted(asset.part_name for asset in model.assets.values()) == [
        "word/media/image1.tmp",
        "word/media/image2.png",
    ]


def test_the_two_parts_get_different_ids(twins):
    model = DocxParser().parse(twins)

    assert len(model.assets) == 2, "one id for two parts means one of them is gone"


def test_they_still_agree_on_the_content_hash(twins):
    digests = {asset.sha256 for asset in DocxParser().parse(twins).assets.values()}

    assert len(digests) == 1, "same bytes, so the same content hash"


def test_both_parts_come_out_the_other_side(tmp_path, twins):
    with ZipFile(_rebuild(tmp_path, twins)) as archive:
        media = sorted(name for name in archive.namelist() if name.startswith("word/media/"))

    assert len(media) == 2, f"a part was dropped: {media}"


def test_g10_counts_the_same_number_of_images(tmp_path, twins):
    before = g10_projection(twins)
    after = g10_projection(_rebuild(tmp_path, twins, "-g10"))

    assert after["part_kinds"].get("image/png") == before["part_kinds"].get("image/png")


# --- the ids that must not move ------------------------------------------------

def test_a_document_without_a_collision_keeps_its_id(single):
    model = DocxParser().parse(single)

    (asset_id,) = model.assets
    assert asset_id == f"asset_{next(iter(model.assets.values())).sha256[:16]}"


def test_the_ids_are_the_same_every_time(twins):
    first = sorted(DocxParser().parse(twins).assets)
    second = sorted(DocxParser().parse(twins).assets)

    assert first == second
