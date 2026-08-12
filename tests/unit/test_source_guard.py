from pathlib import Path
import pytest
from word_replica.services.source_guard import capture_source, assert_source_unchanged
from word_replica.domain.errors import SourceIntegrityError


def test_source_hash_detects_mutation(tmp_path: Path):
    source = tmp_path / "paper.docx"
    source.write_bytes(b"before")
    snapshot = capture_source(source)
    source.write_bytes(b"after")
    with pytest.raises(SourceIntegrityError):
        assert_source_unchanged(snapshot)
