"""The corpus database.

SQLite, for three concrete reasons rather than habit:

* It is stdlib. ``pyarrow`` and ``zstandard`` are not installed here, and on a
  machine with under six gigabytes free a 130 MB dependency is a real cost
  measured against a ~2 GB budget for the whole lab.
* A streaming ingest that may be interrupted needs crash-atomic appends. A WAL
  commit gives that; a torn last line in a JSONL file needs bespoke recovery.
* The scheduler's central operation -- "the N highest-scoring documents in this
  tier that have not run recently" -- is an indexed query. Over JSONL it is a
  full parse of every row, every time the queue is rebuilt.

Two storage decisions worth knowing: identity is the content hash, so
re-ingesting the same bytes is idempotent and free; and scores are written on a
*second* pass, because rarity needs corpus-wide document frequencies that do
not exist until the corpus has been counted.
"""
from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import sqlite3
from typing import Any, Iterator, Mapping

from word_replica.lab.fingerprint import DocxFingerprint, NAMESPACE_BITS
from word_replica.lab.scoring import ComplexityScore, FEATURE_BITS

__all__ = ["SCHEMA_VERSION", "LabStore"]

SCHEMA_VERSION = 1

# Columns carried straight through from DocxFingerprint.to_row(). Kept explicit
# rather than derived so a fingerprint schema change is a deliberate migration
# and not a silent column drift.
_FINGERPRINT_COLUMNS: tuple[str, ...] = (
    "byte_size", "schema_version", "extractor_ok", "extract_error",
    "risk_class", "risk_reasons", "word_open_allowed",
    "part_count", "xml_part_count", "media_count", "total_uncompressed_bytes",
    "max_part_bytes", "compression_ratio", "document_xml_bytes",
    "missing_content_type_count", "missing_rel_target_count", "external_rel_count",
    "rel_type_count", "ns_mask", "unknown_ns_count", "unknown_ns_sample",
    "distinct_qname_count", "total_element_count", "max_xml_depth",
    "mc_alternate_content_count",
    "paragraph_count", "run_count", "char_count", "table_count", "max_table_depth",
    "row_count", "cell_count", "merged_cell_count", "section_count",
    "multicolumn_section_count",
    "drawing_inline_count", "drawing_anchor_count", "behind_text_count",
    "textbox_count", "vml_shape_count", "ole_object_count", "chart_count",
    "chartex_count", "smartart_count", "omml_count",
    "field_count", "pagination_field_count", "bookmark_count", "hyperlink_count",
    "footnote_count", "endnote_count", "comment_count", "comment_reply_count",
    "revision_count", "sdt_count",
    "numbering_def_count", "numbering_max_level", "style_def_count",
    "header_count", "footer_count", "even_odd_headers", "keep_together_count",
    "page_break_before_count", "frame_pr_count", "tbl_layout_auto_count",
    "tab_leader_count", "complex_script_run_count", "rtl_run_count",
    "embedded_font_count", "custom_xml_part_count", "compat_setting_count",
    "compat_mode", "doc_protection", "track_changes_on", "macro_present",
    "activex_count", "model_fingerprint", "model_error",
)

_DDL = f"""
CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS document (
    sha256      TEXT PRIMARY KEY,
    source      TEXT NOT NULL,
    source_ref  TEXT,
    licence     TEXT,
    pinned_path TEXT,
    holdout     INTEGER NOT NULL DEFAULT 0,
    ingested_utc TEXT,
    {", ".join(f"{name} {'REAL' if name == 'compression_ratio' else 'TEXT' if name in
        ('extract_error', 'risk_class', 'risk_reasons', 'unknown_ns_sample',
         'model_fingerprint', 'model_error') else 'INTEGER'}" for name in _FINGERPRINT_COLUMNS)},
    structural_score REAL, rarity_score REAL, failure_score REAL, layout_score REAL,
    complexity_score REAL, score_model_version INTEGER, feature_mask INTEGER
);

CREATE INDEX IF NOT EXISTS document_complexity ON document(complexity_score DESC);
CREATE INDEX IF NOT EXISTS document_risk       ON document(risk_class, holdout);
CREATE INDEX IF NOT EXISTS document_source     ON document(source);
CREATE INDEX IF NOT EXISTS document_ns         ON document(ns_mask);
"""


@dataclass(slots=True)
class LabStore:
    connection: sqlite3.Connection

    @classmethod
    def open(cls, path: Path) -> "LabStore":
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(path)
        connection.row_factory = sqlite3.Row
        # WAL is what makes an interrupted ingest recoverable without bespoke
        # crash handling; NORMAL synchronous is the right trade for a corpus we
        # can always re-derive from its sources.
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=NORMAL")
        connection.execute("PRAGMA foreign_keys=ON")
        store = cls(connection)
        store._apply_schema()
        return store

    def __enter__(self) -> "LabStore":
        return self

    def __exit__(self, *_exc: object) -> None:
        self.close()

    def close(self) -> None:
        self.connection.commit()
        self.connection.close()

    def _apply_schema(self) -> None:
        with self.connection:
            self.connection.executescript(_DDL)
            self.connection.execute(
                "INSERT OR IGNORE INTO meta(key, value) VALUES ('schema_version', ?)",
                (str(SCHEMA_VERSION),),
            )

    def schema_version(self) -> int:
        row = self.connection.execute("SELECT value FROM meta WHERE key='schema_version'").fetchone()
        return int(row["value"]) if row else 0

    @contextmanager
    def transaction(self) -> Iterator[None]:
        """A batch that either lands whole or not at all.

        The leading commit is load-bearing. Python's sqlite3 opens an implicit
        transaction on the first statement and holds it open, so a rollback here
        would otherwise discard every uncommitted write made *before* this
        block -- turning a crash in one ingest batch into the loss of the
        previous one. Committing first makes the rollback boundary exactly this
        block and nothing earlier.
        """
        self.connection.commit()
        try:
            yield
        except BaseException:
            self.connection.rollback()
            raise
        else:
            self.connection.commit()

    # --- writes ---------------------------------------------------------

    def upsert(
        self,
        fingerprint: DocxFingerprint,
        *,
        source: str,
        source_ref: str | None = None,
        licence: str | None = None,
        pinned_path: str | None = None,
        ingested_utc: str | None = None,
    ) -> None:
        """Identity is the content hash, so re-ingesting the same bytes is free.

        Scores are deliberately not written here -- see write_score.
        """
        row = fingerprint.to_row()
        values: dict[str, Any] = {name: row.get(name) for name in _FINGERPRINT_COLUMNS}
        values.update(
            sha256=fingerprint.sha256,
            source=source,
            source_ref=source_ref,
            licence=licence,
            pinned_path=pinned_path,
            ingested_utc=ingested_utc,
        )
        columns = list(values)
        assignments = ", ".join(f"{name}=excluded.{name}" for name in columns if name != "sha256")
        self.connection.execute(
            f"INSERT INTO document ({', '.join(columns)}) VALUES ({', '.join('?' for _ in columns)}) "
            f"ON CONFLICT(sha256) DO UPDATE SET {assignments}",
            [values[name] for name in columns],
        )

    def write_score(self, digest: str, score: ComplexityScore) -> None:
        self.connection.execute(
            "UPDATE document SET structural_score=?, rarity_score=?, failure_score=?, "
            "layout_score=?, complexity_score=?, score_model_version=?, feature_mask=? "
            "WHERE sha256=?",
            (
                score.structural, score.rarity, score.failure, score.layout,
                score.total, score.model_version, score.feature_mask, digest,
            ),
        )

    def write_feature_mask(self, digest: str, mask: int) -> None:
        self.connection.execute("UPDATE document SET feature_mask=? WHERE sha256=?", (mask, digest))

    def assign_holdout(self, digest: str, *, percent: int = 5) -> bool:
        """Decide holdout membership from the content hash alone.

        Never from a score, and never after a failure has been seen -- otherwise
        the holdout number stops meaning what it claims to mean.
        """
        member = int(sha256((digest + "holdout").encode("utf-8")).hexdigest()[:8], 16) % 100 < percent
        self.connection.execute("UPDATE document SET holdout=? WHERE sha256=?", (int(member), digest))
        return member

    # --- reads ----------------------------------------------------------

    def count_documents(self) -> int:
        return int(self.connection.execute("SELECT COUNT(*) AS n FROM document").fetchone()["n"])

    def get(self, digest: str) -> sqlite3.Row | None:
        return self.connection.execute("SELECT * FROM document WHERE sha256=?", (digest,)).fetchone()

    def iter_documents(self, *, only_unscored: bool = False) -> Iterator[sqlite3.Row]:
        query = "SELECT * FROM document"
        if only_unscored:
            query += " WHERE complexity_score IS NULL"
        yield from self.connection.execute(query)

    def document_frequency(self) -> dict[int, int]:
        """How many documents carry each feature -- the input to rarity.

        One pass over the stored masks; this is the second-pass step that makes
        rarity meaningful instead of uniformly maximal.
        """
        frequency = {index: 0 for index in range(len(FEATURE_BITS))}
        for (mask,) in self.connection.execute(
            "SELECT feature_mask FROM document WHERE feature_mask IS NOT NULL"
        ):
            for index in range(len(FEATURE_BITS)):
                if mask >> index & 1:
                    frequency[index] += 1
        return frequency

    def rank(
        self,
        *,
        limit: int = 50,
        exclude_risk: tuple[str, ...] = ("HOSTILE",),
        include_holdout: bool = False,
    ) -> list[sqlite3.Row]:
        clauses = ["complexity_score IS NOT NULL"]
        params: list[Any] = []
        if exclude_risk:
            clauses.append(f"risk_class NOT IN ({', '.join('?' for _ in exclude_risk)})")
            params.extend(exclude_risk)
        if not include_holdout:
            clauses.append("holdout = 0")
        params.append(limit)
        return list(
            self.connection.execute(
                f"SELECT * FROM document WHERE {' AND '.join(clauses)} "
                f"ORDER BY complexity_score DESC, sha256 ASC LIMIT ?",
                params,
            )
        )

    def namespace_histogram(self) -> dict[str, int]:
        """Documents per namespace -- the empirical case for the G10 gate.

        The interesting rows are w15/w16/cx/asvg/strict: content Word 2010
        ignores or cannot write, and which every existing gate is blind to.
        """
        histogram = {name: 0 for name in NAMESPACE_BITS}
        for (mask,) in self.connection.execute(
            "SELECT ns_mask FROM document WHERE ns_mask IS NOT NULL"
        ):
            for index, name in enumerate(NAMESPACE_BITS):
                if mask >> index & 1:
                    histogram[name] += 1
        return {name: count for name, count in histogram.items() if count}

    def risk_breakdown(self) -> dict[str, int]:
        return {
            row["risk_class"]: row["n"]
            for row in self.connection.execute(
                "SELECT risk_class, COUNT(*) AS n FROM document GROUP BY risk_class"
            )
        }

    def parser_failures(self, limit: int = 50) -> list[sqlite3.Row]:
        """Documents our own parser could not model -- the M1 headline number."""
        return list(
            self.connection.execute(
                "SELECT sha256, source, source_ref, model_error, extract_error, risk_class "
                "FROM document WHERE model_error IS NOT NULL OR extractor_ok = 0 "
                "ORDER BY source, source_ref LIMIT ?",
                (limit,),
            )
        )


def scoring_context(store: LabStore, failure_stats: Mapping[int, tuple[int, int]] | None = None):
    """Build the second-pass context from what the corpus now contains."""
    from word_replica.lab.scoring import ScoringContext

    return ScoringContext(
        corpus_n=store.count_documents(),
        document_frequency=store.document_frequency(),
        failure_stats=dict(failure_stats or {}),
    )
