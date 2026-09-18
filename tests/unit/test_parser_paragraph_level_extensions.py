"""Content a paragraph holds outside a run is still content.

`parse_paragraph` dispatches on the paragraph's direct children: properties,
runs, bookmarks, and the wrappers it unwraps to reach runs. An equation is none
of those. `m:oMath` sits directly in `w:p`, beside the runs rather than inside
one, so it matched nothing and was dropped -- and because the renderer builds
the paragraph from the model, dropped meant gone.

Two corpus documents are affected, math-literal.docx and math-accents.docx, and
G10 reports the officeDocument/2006/math namespace disappearing. The equation
does not come back as a broken equation; it does not come back at all.

Only children outside the `w:` namespace are carried, the same line drawn for
run properties: an unmodelled `w:` element is a question about the model's
coverage, and copying one in beside content the renderer also writes risks
emitting it twice.

Position is kept, not just presence. An equation between two sentences belongs
between them, so each fragment records how many runs preceded it and goes back
in the same place.
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
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

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

EQUATION = (
    "<m:oMath>"
    "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
    "</m:oMath>"
)


def _run(text):
    return f"<w:r><w:t>{text}</w:t></w:r>"


def _write(path, body):
    document = (
        f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:m="{M}">'
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


def _paragraph(path):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml")).find(f".//{{{W}}}p")


def _child_names(paragraph):
    return [child.tag.rsplit("}", 1)[-1] for child in paragraph]


# --- the corpus shape: a run, then the equation --------------------------------

TRAILING = f"<w:p>{_run('')}{EQUATION}</w:p>"


@pytest.fixture
def trailing(tmp_path):
    return _write(tmp_path / "trailing.docx", TRAILING)


def test_the_equation_survives(tmp_path, trailing):
    assert _paragraph(_rebuild(tmp_path, trailing)).find(f"{{{M}}}oMath") is not None


def test_its_contents_survive_intact(tmp_path, trailing):
    equation = _paragraph(_rebuild(tmp_path, trailing, "-inner")).find(f"{{{M}}}oMath")

    assert [node.text for node in equation.iter(f"{{{M}}}t")] == ["x", "2"]


def test_g10_keeps_the_maths_namespace(tmp_path, trailing):
    after = g10_projection(_rebuild(tmp_path, trailing, "-g10"))

    assert M in after["namespaces"]


# --- position --------------------------------------------------------------

def test_an_equation_between_two_runs_stays_between_them(tmp_path):
    source = _write(
        tmp_path / "between.docx", f"<w:p>{_run('before ')}{EQUATION}{_run(' after')}</w:p>"
    )

    names = _child_names(_paragraph(_rebuild(tmp_path, source, "-between")))

    assert names == ["r", "oMath", "r"]


def test_an_equation_before_every_run_stays_first(tmp_path):
    source = _write(tmp_path / "leading.docx", f"<w:p>{EQUATION}{_run('after')}</w:p>")

    names = _child_names(_paragraph(_rebuild(tmp_path, source, "-leading")))

    assert names == ["oMath", "r"]


def test_two_equations_keep_their_order(tmp_path):
    second = EQUATION.replace("<m:t>x</m:t>", "<m:t>y</m:t>")
    source = _write(tmp_path / "two.docx", f"<w:p>{EQUATION}{_run(' and ')}{second}</w:p>")

    paragraph = _paragraph(_rebuild(tmp_path, source, "-two"))
    equations = paragraph.findall(f"{{{M}}}oMath")

    assert _child_names(paragraph) == ["oMath", "r", "oMath"]
    assert [next(node.iter(f"{{{M}}}t")).text for node in equations] == ["x", "y"]


# --- what must not change ------------------------------------------------------

def test_the_surrounding_text_is_untouched(tmp_path):
    source = _write(
        tmp_path / "text.docx", f"<w:p>{_run('before ')}{EQUATION}{_run(' after')}</w:p>"
    )

    output = _rebuild(tmp_path, source, "-text")

    assert DocxParser().parse(output).plain_text().strip() == "before  after"


def test_a_paragraph_with_no_extension_gains_none(tmp_path):
    source = _write(tmp_path / "plain.docx", f"<w:p>{_run('plain')}</w:p>")

    assert _child_names(_paragraph(_rebuild(tmp_path, source, "-plain"))) == ["r"]
