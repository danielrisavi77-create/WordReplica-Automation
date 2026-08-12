from pathlib import Path
import pytest
from tests.fixtures.build_fixtures import build_all_corpus


@pytest.fixture(scope="session")
def corpus_dir(tmp_path_factory) -> Path:
    path = tmp_path_factory.mktemp("word_replica_corpus")
    build_all_corpus(path)
    return path
