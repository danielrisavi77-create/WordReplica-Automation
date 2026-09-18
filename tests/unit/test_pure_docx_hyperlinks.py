"""Hyperlinks must survive as links, not as the words they were made of.

Found only once the corpus sample grew from 60 documents to 400 -- five
documents whose external hyperlink relationships were present in the source and
absent from the rebuild. The small sample had none, which is the argument for
running the lab wider rather than deeper.

`parse_paragraph` unwraps `w:hyperlink` to reach the runs inside it, so the
link text survived and the link did not: every hyperlink in every document came
out as plain text. G0 compares body text and saw nothing wrong, because the
text is exactly what does survive.
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
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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
    f'<Relationship Id="rId1" Target="word/document.xml" Type="{REL}/officeDocument"/>'
    "</Relationships>"
).encode("utf-8")


def _source(path, *, body: str, doc_rels: str = ""):
    rels = (
        '<?xml version="1.0"?><Relationships '
        'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{doc_rels}</Relationships>"
    ).encode("utf-8")
    document = (
        '<?xml version="1.0"?><w:document '
        f'xmlns:w="{W}" xmlns:r="{REL}">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    ).encode("utf-8")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", rels)
        archive.writestr("word/document.xml", document)
    return path


EXTERNAL_REL = (
    f'<Relationship Id="rId7" TargetMode="External" Type="{REL}/hyperlink" '
    'Target="https://example.com/docs"/>'
)
LINKED = (
    '<w:p><w:r><w:t>see </w:t></w:r>'
    '<w:hyperlink r:id="rId7"><w:r><w:t>the manual</w:t></w:r></w:hyperlink>'
    '<w:r><w:t> for more</w:t></w:r></w:p>'
)


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


def _rels(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/_rels/document.xml.rels"))


@pytest.fixture
def linked(tmp_path):
    return _source(tmp_path / "link.docx", body=LINKED, doc_rels=EXTERNAL_REL)


# --- capture ------------------------------------------------------------------

def test_the_parser_records_which_runs_were_inside_a_link(linked):
    model = DocxParser().parse(linked)

    linked_runs = [
        run
        for paragraph in model.iter_paragraphs()
        for run in paragraph.runs
        if run.properties.get("hyperlink")
    ]
    assert [run.text for run in linked_runs] == ["the manual"]
    assert linked_runs[0].properties["hyperlink"]["target"] == "https://example.com/docs"


def test_text_outside_the_link_is_not_marked(linked):
    model = DocxParser().parse(linked)
    runs = [run for paragraph in model.iter_paragraphs() for run in paragraph.runs]

    assert [bool(run.properties.get("hyperlink")) for run in runs] == [False, True, False]


# --- re-emission --------------------------------------------------------------

def test_the_link_comes_back_as_a_link(tmp_path, linked):
    document = _document(_rebuild(tmp_path, linked))

    hyperlinks = document.findall(f".//{{{W}}}hyperlink")
    assert len(hyperlinks) == 1
    assert hyperlinks[0].findall(f".//{{{W}}}t")[0].text == "the manual"


def test_the_link_target_survives(tmp_path, linked):
    output = _rebuild(tmp_path, linked, "-target")
    rels = _rels(output)

    targets = {
        node.get("Target")
        for node in rels
        if (node.get("Type") or "").endswith("/hyperlink")
    }
    assert targets == {"https://example.com/docs"}
    assert all(
        node.get("TargetMode") == "External"
        for node in rels
        if (node.get("Type") or "").endswith("/hyperlink")
    )


def test_the_link_points_at_a_relationship_that_exists(tmp_path, linked):
    output = _rebuild(tmp_path, linked, "-rel")
    document, rels = _document(output), _rels(output)

    used = {
        node.get(f"{{{REL}}}id")
        for node in document.findall(f".//{{{W}}}hyperlink")
        if node.get(f"{{{REL}}}id")
    }
    declared = {node.get("Id") for node in rels}
    assert used and used <= declared


def test_the_surrounding_text_keeps_its_order(tmp_path, linked):
    assert DocxParser().parse(_rebuild(tmp_path, linked, "-order")).plain_text().strip() == (
        "see the manual for more"
    )


def test_g10_no_longer_reports_the_link_as_lost(tmp_path, linked):
    before = g10_projection(linked)
    after = g10_projection(_rebuild(tmp_path, linked, "-g10"))

    assert after["relationship_graph"].get("hyperlink") == before["relationship_graph"].get("hyperlink")
    assert after["dangling_relationships"] == []


# --- the other kind -----------------------------------------------------------

def test_an_internal_anchor_link_survives(tmp_path):
    # A link to a bookmark inside the same document has no relationship at all,
    # only a w:anchor attribute.
    source = _source(
        tmp_path / "anchor.docx",
        body='<w:p><w:hyperlink w:anchor="Chapter2"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>',
    )

    document = _document(_rebuild(tmp_path, source, "-anchor"))

    hyperlink = document.find(f".//{{{W}}}hyperlink")
    assert hyperlink is not None
    assert hyperlink.get(f"{{{W}}}anchor") == "Chapter2"


def test_a_document_without_links_gains_none(tmp_path):
    source = _source(tmp_path / "plain.docx", body="<w:p><w:r><w:t>plain</w:t></w:r></w:p>")

    assert _document(_rebuild(tmp_path, source, "-plain")).find(f".//{{{W}}}hyperlink") is None
