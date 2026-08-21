"""A rebuild can keep its project somewhere other than beside the source.

`ProjectStore.create_project` puts the project in `source.parent /
f"{source.stem}_rebuild"`, regardless of `app_root` -- which only ever held the
projects index. For someone rebuilding their own document that is the right
place: the project lands next to the thing it came from.

For the fidelity lab it is wrong. The lab reads a corpus it treats as
read-only, so a 400-document pass wrote 400 full projects -- source snapshot,
output, logs, QA reports -- into the corpus tree itself. Measured: 469 leftover
`*_rebuild` directories holding 263 MB, 83% of everything under lab/, and they
accumulate again on every pass. The lane already passes a temporary directory
as `app_root`; it just had no effect on where projects went.

So placement becomes a choice the caller makes, and the default is exactly what
it was. The corpus tree a reader hands us is not ours to write into.
"""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from word_replica.config import RebuildOptions
from word_replica.domain.enums import RendererChoice
from word_replica.services.project_store import ProjectStore

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
DOCUMENT = (
    f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>'
    "<w:p><w:r><w:t>hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
).encode("utf-8")


@pytest.fixture
def corpus(tmp_path):
    """A source document in a tree the caller does not own."""
    directory = tmp_path / "corpus"
    directory.mkdir()
    path = directory / "sample.docx"
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr(
            "word/_rels/document.xml.rels",
            b'<?xml version="1.0"?><Relationships '
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        )
        archive.writestr("word/document.xml", DOCUMENT)
    return path


def _options():
    return RebuildOptions(renderer=RendererChoice.DOCX)


def _rebuild_dirs(directory: Path):
    return sorted(p.name for p in directory.iterdir() if p.name.endswith("_rebuild"))


def test_by_default_the_project_still_lands_beside_the_source(tmp_path, corpus):
    store = ProjectStore(app_root=tmp_path / "app")

    paths = store.create_project(corpus, _options())

    assert Path(paths.root).parent.parent == corpus.parent
    assert _rebuild_dirs(corpus.parent) == ["sample_rebuild"]


def test_the_caller_can_keep_projects_under_the_app_root(tmp_path, corpus):
    app_root = tmp_path / "app"
    store = ProjectStore(app_root=app_root, projects_under_app_root=True)

    paths = store.create_project(corpus, _options())

    assert app_root in Path(paths.root).parents


def test_then_nothing_is_written_beside_the_source(tmp_path, corpus):
    store = ProjectStore(app_root=tmp_path / "app", projects_under_app_root=True)

    store.create_project(corpus, _options())

    assert _rebuild_dirs(corpus.parent) == []


def test_the_source_document_itself_is_untouched(tmp_path, corpus):
    before = corpus.read_bytes()
    store = ProjectStore(app_root=tmp_path / "app", projects_under_app_root=True)

    store.create_project(corpus, _options())

    assert corpus.read_bytes() == before


def test_the_project_still_has_everything_it_needs(tmp_path, corpus):
    store = ProjectStore(app_root=tmp_path / "app", projects_under_app_root=True)

    paths = store.create_project(corpus, _options())

    root = Path(paths.root)
    assert (root / "project.json").is_file()
    assert (root / "source_snapshot" / corpus.name).is_file()
    for name in ("working", "output", "backups", "logs", "qa"):
        assert (root / name).is_dir(), name


def test_two_documents_of_the_same_name_do_not_collide(tmp_path, corpus):
    other_dir = tmp_path / "corpus2"
    other_dir.mkdir()
    other = other_dir / corpus.name
    other.write_bytes(corpus.read_bytes())
    store = ProjectStore(app_root=tmp_path / "app", projects_under_app_root=True)

    first = Path(store.create_project(corpus, _options()).root)
    second = Path(store.create_project(other, _options()).root)

    assert first != second
