from pathlib import Path

import pytest

from scripts.persistent_harness.migrate import MigrationCollisionError, import_docx_corpus


def test_migration_copies_without_moving_source(tmp_path):
    source = tmp_path / "old"
    dest = tmp_path / "new"
    source.mkdir(); dest.mkdir()
    original = source / "a.docx"
    original.write_bytes(b"same")
    result = import_docx_corpus(source, dest)
    assert result.copied == 1
    assert original.exists()
    assert (dest / "a.docx").read_bytes() == b"same"


def test_migration_refuses_different_same_name_file(tmp_path):
    source = tmp_path / "old"
    dest = tmp_path / "new"
    source.mkdir(); dest.mkdir()
    (source / "a.docx").write_bytes(b"source")
    (dest / "a.docx").write_bytes(b"different")
    with pytest.raises(MigrationCollisionError, match="different SHA-256"):
        import_docx_corpus(source, dest)
    assert (dest / "a.docx").read_bytes() == b"different"


def test_migration_skips_identical_existing_file(tmp_path):
    source = tmp_path / "old"
    dest = tmp_path / "new"
    source.mkdir(); dest.mkdir()
    (source / "a.docx").write_bytes(b"same")
    (dest / "a.docx").write_bytes(b"same")
    result = import_docx_corpus(source, dest)
    assert result.copied == 0
    assert result.skipped_identical == 1
