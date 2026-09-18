"""The corpus database.

SQLite because it is stdlib (pyarrow and zstandard are not installed, and on a
machine with single-digit gigabytes free a 130 MB dependency is a real cost),
because a streaming ingest that may be interrupted needs crash-atomic appends,
and because the scheduler's central operation -- "the N highest-scoring
documents in this tier that have not run recently" -- is an indexed query, not
a scan.
"""
import sqlite3

import pytest

from word_replica.lab.fingerprint import DocxFingerprint
from word_replica.lab.scoring import ScoringContext, score_document
from word_replica.lab.store import LabStore, SCHEMA_VERSION


def _fp(digest: str, **kwargs) -> DocxFingerprint:
    base = dict(sha256=digest, byte_size=1000, extractor_ok=True, risk_class="VALID")
    return DocxFingerprint(**{**base, **kwargs})


EMPTY = ScoringContext(corpus_n=0, document_frequency={}, failure_stats={})


@pytest.fixture
def store(tmp_path):
    with LabStore.open(tmp_path / "lab.db") as opened:
        yield opened


def test_schema_applies_to_a_fresh_database(store):
    assert store.schema_version() == SCHEMA_VERSION
    assert store.count_documents() == 0


def test_reopening_an_existing_database_does_not_reapply_the_schema(tmp_path):
    path = tmp_path / "lab.db"
    with LabStore.open(path) as first:
        first.upsert(_fp("a" * 64), source="test", source_ref="one")
    with LabStore.open(path) as second:
        assert second.count_documents() == 1
        assert second.schema_version() == SCHEMA_VERSION


def test_a_document_round_trips(store):
    fingerprint = _fp("b" * 64, table_count=3, max_table_depth=2, ns_mask=7)

    store.upsert(fingerprint, source="libreoffice", source_ref="ooxmlexport/data/x.docx")

    row = store.get("b" * 64)
    assert row["table_count"] == 3
    assert row["max_table_depth"] == 2
    assert row["ns_mask"] == 7
    assert row["source"] == "libreoffice"
    assert row["source_ref"] == "ooxmlexport/data/x.docx"


def test_reingesting_the_same_bytes_is_idempotent(store):
    store.upsert(_fp("c" * 64, table_count=1), source="a", source_ref="one")
    store.upsert(_fp("c" * 64, table_count=9), source="a", source_ref="one")

    assert store.count_documents() == 1
    assert store.get("c" * 64)["table_count"] == 9


def test_scores_are_written_on_a_second_pass(store):
    fingerprint = _fp("d" * 64, table_count=2)
    store.upsert(fingerprint, source="a", source_ref="one")
    assert store.get("d" * 64)["complexity_score"] is None

    store.write_score("d" * 64, score_document(fingerprint, EMPTY))

    row = store.get("d" * 64)
    assert row["complexity_score"] is not None
    assert row["structural_score"] is not None
    assert row["score_model_version"] == 1


def test_document_frequency_is_one_group_by_over_the_corpus(store):
    from word_replica.lab.scoring import FEATURE_BITS, feature_mask

    store.upsert(_fp("e" * 64, table_count=1), source="a", source_ref="1")
    store.upsert(_fp("f" * 64, table_count=1, chartex_count=1), source="a", source_ref="2")
    store.upsert(_fp("0" * 64), source="a", source_ref="3")
    for digest, fingerprint in (
        ("e" * 64, _fp("e" * 64, table_count=1)),
        ("f" * 64, _fp("f" * 64, table_count=1, chartex_count=1)),
        ("0" * 64, _fp("0" * 64)),
    ):
        store.write_feature_mask(digest, feature_mask(fingerprint))

    frequency = store.document_frequency()

    assert frequency[FEATURE_BITS.index("tables")] == 2
    assert frequency[FEATURE_BITS.index("chartex")] == 1
    assert frequency.get(FEATURE_BITS.index("smartart"), 0) == 0


def test_ranking_orders_by_complexity_and_excludes_hostile_documents(store):
    for index, (digest, tables, risk) in enumerate([
        ("1" * 64, 1, "VALID"),
        ("2" * 64, 9, "VALID"),
        ("3" * 64, 20, "HOSTILE"),
    ]):
        fingerprint = _fp(digest, table_count=tables, max_table_depth=2, risk_class=risk)
        store.upsert(fingerprint, source="a", source_ref=str(index))
        store.write_score(digest, score_document(fingerprint, EMPTY))

    ranked = store.rank(limit=10, exclude_risk=("HOSTILE",))

    assert [row["sha256"] for row in ranked] == ["2" * 64, "1" * 64]


def test_namespace_histogram_counts_documents_not_elements(store):
    store.upsert(_fp("4" * 64, ns_mask=0b101), source="a", source_ref="1")
    store.upsert(_fp("5" * 64, ns_mask=0b100), source="a", source_ref="2")

    histogram = store.namespace_histogram()

    assert histogram["w"] == 1      # bit 0, only the first document
    assert histogram["w15"] == 2    # bit 2, both


def test_risk_breakdown_is_reported(store):
    store.upsert(_fp("6" * 64, risk_class="VALID"), source="a", source_ref="1")
    store.upsert(_fp("7" * 64, risk_class="RECOVERABLE"), source="a", source_ref="2")
    store.upsert(_fp("8" * 64, risk_class="RECOVERABLE"), source="a", source_ref="3")

    assert store.risk_breakdown() == {"VALID": 1, "RECOVERABLE": 2}


def test_holdout_membership_is_stable_for_the_same_document(store):
    first = store.assign_holdout("9" * 64)
    second = store.assign_holdout("9" * 64)

    assert first is second
    # And it is decided by the content hash alone, never by a score, so it
    # cannot be influenced by anything the lab learns later.
    assert store.assign_holdout("9" * 64, percent=0) is False
    assert store.assign_holdout("9" * 64, percent=100) is True


def test_an_interrupted_write_leaves_the_database_readable(tmp_path):
    path = tmp_path / "lab.db"
    with LabStore.open(path) as store:
        store.upsert(_fp("a" * 64), source="a", source_ref="1")
        with pytest.raises(RuntimeError):
            with store.transaction():
                store.upsert(_fp("b" * 64), source="a", source_ref="2")
                raise RuntimeError("crash mid-batch")
        assert store.count_documents() == 1

    with sqlite3.connect(path) as connection:
        assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
