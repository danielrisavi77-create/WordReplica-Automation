"""Run properties from outside the main namespace are still run properties.

`parse_run` reads the run properties the model represents -- fonts, size,
colour, bold, language and the rest -- from `w:rPr`. Anything else in there was
ignored, and since the renderer builds a fresh `w:rPr` from the model, ignored
meant gone.

What is in there, in practice, is Word's own typography extensions:
`w14:textFill` with its gradient stops, `w14:shadow`, `w14:numForm`. Two corpus
documents lose them -- TextEffects_TextFill.docx from the body and
SdtContent.docx from a footer -- and G10 reports the
office/word/2010/wordml namespace disappearing. A run that was a gradient comes
back flat.

Only children outside the `w:` namespace are carried. An unmodelled `w:`
property is a different question: the renderer writes that namespace itself, in
a required order, and copying elements into the middle of that is how a document
starts needing repair. Extension elements sit at the end, which is where Word
puts them and where they are appended here.
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
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument'
    '.wordprocessingml.document.main+xml"/>'
    "</Types>"
).encode("utf-8")
ROOT_RELS = (
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Target="word/document.xml" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")

TEXT_FILL = (
    "<w14:textFill><w14:gradFill><w14:gsLst>"
    '<w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>'
    '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>'
    '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/>'
    "</w14:gradFill></w14:textFill>"
)

WITH_EFFECTS = (
    "<w:p><w:r><w:rPr>"
    '<w:b/><w:color w:val="112233"/><w:sz w:val="28"/>'
    f"{TEXT_FILL}"
    "</w:rPr><w:t>gradient text</w:t></w:r></w:p>"
)
EFFECTS_ONLY = f"<w:p><w:r><w:rPr>{TEXT_FILL}</w:rPr><w:t>only effects</w:t></w:r></w:p>"
PLAIN = "<w:p><w:r><w:t>plain</w:t></w:r></w:p>"


def _write(path, body):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:w14="{W14}" '
        f'xmlns:mc="{MC}" mc:Ignorable="w14">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", document)
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


def _run_properties(path):
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    return root.find(f".//{{{W}}}r/{{{W}}}rPr")


@pytest.fixture
def effects(tmp_path):
    return _write(tmp_path / "effects.docx", WITH_EFFECTS)


def test_the_text_fill_survives(tmp_path, effects):
    r_pr = _run_properties(_rebuild(tmp_path, effects))

    assert r_pr is not None
    assert r_pr.find(f"{{{W14}}}textFill") is not None


def test_the_gradient_stops_survive_intact(tmp_path, effects):
    r_pr = _run_properties(_rebuild(tmp_path, effects, "-stops"))

    stops = r_pr.findall(f".//{{{W14}}}gs/{{{W14}}}srgbClr")
    assert [stop.get(f"{{{W14}}}val") for stop in stops] == ["FF0000", "0000FF"]


def test_the_modelled_properties_are_unchanged(tmp_path, effects):
    r_pr = _run_properties(_rebuild(tmp_path, effects, "-modelled"))

    assert r_pr.find(f"{{{W}}}b") is not None
    assert r_pr.find(f"{{{W}}}color").get(f"{{{W}}}val") == "112233"
    assert r_pr.find(f"{{{W}}}sz").get(f"{{{W}}}val") == "28"


def test_the_extension_comes_after_the_modelled_properties(tmp_path, effects):
    r_pr = _run_properties(_rebuild(tmp_path, effects, "-order"))
    names = [child.tag for child in r_pr]

    assert names[-1] == f"{{{W14}}}textFill"


def test_g10_keeps_the_namespace(tmp_path, effects):
    after = g10_projection(_rebuild(tmp_path, effects, "-g10"))

    assert W14 in after["namespaces"]


def test_a_run_with_only_extension_properties_still_gets_them(tmp_path):
    source = _write(tmp_path / "only.docx", EFFECTS_ONLY)

    r_pr = _run_properties(_rebuild(tmp_path, source, "-only"))

    assert r_pr is not None, "no w:rPr was written at all, so the effect had nowhere to go"
    assert r_pr.find(f"{{{W14}}}textFill") is not None


# --- what must not change ------------------------------------------------------

def test_a_run_without_properties_gains_none(tmp_path):
    source = _write(tmp_path / "plain.docx", PLAIN)

    assert _run_properties(_rebuild(tmp_path, source, "-plain")) is None


def test_the_text_is_untouched(tmp_path, effects):
    from word_replica.parser.parser import DocxParser

    output = _rebuild(tmp_path, effects, "-text")

    assert DocxParser().parse(output).plain_text().strip() == "gradient text"
