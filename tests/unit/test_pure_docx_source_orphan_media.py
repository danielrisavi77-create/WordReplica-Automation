from zipfile import ZIP_DEFLATED, ZipFile

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
)
from word_replica.services.rebuild import RebuildService


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _source_with_unreachable_media(path):
    content_types = (
        '<?xml version="1.0"?><Types '
        'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.document.main+xml"/>'
        "</Types>"
    ).encode("utf-8")
    root_rels = (
        f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Target="word/document.xml" '
        f'Type="{REL}/officeDocument"/>'
        "</Relationships>"
    ).encode("utf-8")
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
        "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/>"
        "</w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr(
            "word/_rels/document.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL}"/>'.encode(
                "utf-8"
            ),
        )
        archive.writestr("word/document.xml", document)
        archive.writestr("word/media/orphan.png", b"source-orphan-image")
    return path


def test_source_unreachable_media_is_preserved_without_renderer_loss_warning(tmp_path):
    source = _source_with_unreachable_media(tmp_path / "source.docx")
    result = RebuildService(app_root=tmp_path / "app").rebuild(
        source,
        RebuildOptions(
            renderer=RendererChoice.DOCX,
            fidelity=FidelityMode.FULL,
            metadata=MetadataMode.PRESERVE,
            reconstruction_mode=ReconstructionMode.INSTANT,
        ),
    )

    assert result.output_path is not None, result.reasons
    assert "PURE_DOCX_ASSET_UNREFERENCED" not in {w.code for w in result.warnings}
    with ZipFile(result.output_path) as archive:
        assert archive.read("word/media/orphan.png") == b"source-orphan-image"
