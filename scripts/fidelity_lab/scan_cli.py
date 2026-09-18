"""Scan a tree of .docx files into the corpus database and report on it.

This is the Milestone 1 loop end to end, and it needs no Word, no network and
no temporary files:

    fingerprint (parallel) -> store -> count features -> score -> rank

The scoring is deliberately two-pass. Rarity is an inverse document frequency,
and with no corpus every feature looks maximally rare, which would rank every
document at 1.0. So the first pass writes rows with rarity NULL, and the second
pass fills it in once the corpus has actually been counted.

    python -m scripts.fidelity_lab.scan_cli --root <dir> --db <path> --top 25
"""
from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys
import time

if __package__ in (None, ""):  # allow `python scripts/fidelity_lab/scan_cli.py`
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from word_replica.lab.fingerprint import extract_fingerprint, namespaces_from_mask
from word_replica.lab.scoring import feature_mask, features_from_mask, score_document
from word_replica.lab.store import LabStore, scoring_context

# Namespaces Microsoft Word 2010 either ignores entirely or cannot write. A
# document carrying one of these is a document whose fidelity the local Word
# oracle cannot honestly judge -- and which every existing G0-G9 gate is blind
# to, since they compare parsed models and the parser does not model it either.
_BEYOND_WORD_2010 = ("w15", "w16", "w16cid", "w16se", "cx", "asvg", "strict")


def _fingerprint_one(path_and_flag: tuple[str, bool]) -> tuple[str, object]:
    path, parse_model = path_and_flag
    return path, extract_fingerprint(Path(path), parse_model=parse_model)


def scan(
    root: Path,
    db_path: Path,
    *,
    source: str,
    workers: int,
    parse_model: bool,
    limit: int | None,
) -> dict:
    paths = sorted(str(p) for p in Path(root).rglob("*.docx"))
    if limit:
        paths = paths[:limit]
    if not paths:
        raise SystemExit(f"no .docx files under {root}")

    started = time.perf_counter()
    now = datetime.now(timezone.utc).isoformat()
    fingerprints: dict[str, object] = {}

    with LabStore.open(db_path) as store:
        with store.transaction():
            if workers > 1:
                with ProcessPoolExecutor(max_workers=workers) as pool:
                    results = pool.map(_fingerprint_one, ((p, parse_model) for p in paths), chunksize=4)
                    for path, fingerprint in results:
                        fingerprints[path] = fingerprint
                        store.upsert(
                            fingerprint, source=source,
                            source_ref=str(Path(path).relative_to(root)), ingested_utc=now,
                        )
                        store.write_feature_mask(fingerprint.sha256, feature_mask(fingerprint))
            else:
                for path in paths:
                    _, fingerprint = _fingerprint_one((path, parse_model))
                    fingerprints[path] = fingerprint
                    store.upsert(
                        fingerprint, source=source,
                        source_ref=str(Path(path).relative_to(root)), ingested_utc=now,
                    )
                    store.write_feature_mask(fingerprint.sha256, feature_mask(fingerprint))
        ingest_seconds = time.perf_counter() - started

        # Second pass: rarity only means something once the corpus is counted.
        context = scoring_context(store)
        with store.transaction():
            for fingerprint in fingerprints.values():
                store.write_score(fingerprint.sha256, score_document(fingerprint, context))
                store.assign_holdout(fingerprint.sha256)

        return {
            "documents": len(paths),
            "ingest_seconds": round(ingest_seconds, 2),
            "ms_per_document": round(ingest_seconds * 1000 / len(paths), 1),
            "elements": sum(f.total_element_count for f in fingerprints.values()),
            "risk": store.risk_breakdown(),
            "namespaces": store.namespace_histogram(),
            "corpus_n": store.count_documents(),
            # WAL mode keeps recent writes in a sidecar until checkpoint, so
            # the main file alone understates what the corpus costs on disk.
            "db_bytes": sum(
                p.stat().st_size
                for p in (db_path, db_path.with_name(db_path.name + "-wal"))
                if p.exists()
            ),
            "holdout": sum(
                1 for row in store.iter_documents() if row["holdout"]
            ),
            "top": [dict(row) for row in store.rank(limit=25)],
            "parser_failures": [dict(row) for row in store.parser_failures(limit=25)],
        }


def _print_report(summary: dict, top: int) -> None:
    print(f"\n{'=' * 78}")
    print(f"  {summary['documents']} documents  |  {summary['elements']:,} XML elements  |  "
          f"{summary['ingest_seconds']}s  ({summary['ms_per_document']} ms/doc)")
    print(f"  corpus now {summary['corpus_n']} documents, database {summary['db_bytes'] / 1024:.0f} KB")
    print("=" * 78)

    print("\nRISK")
    for name, count in sorted(summary["risk"].items()):
        print(f"  {name:<14} {count:>6}")

    print("\nNAMESPACES (documents carrying each)")
    histogram = summary["namespaces"]
    beyond = {name: histogram.get(name, 0) for name in _BEYOND_WORD_2010 if histogram.get(name)}
    for name, count in sorted(histogram.items(), key=lambda item: -item[1]):
        marker = "  <-- Word 2010 cannot judge this" if name in _BEYOND_WORD_2010 else ""
        print(f"  {name:<10} {count:>6}{marker}")
    if beyond:
        total = summary["documents"]
        affected = max(beyond.values())
        print(f"\n  {affected}/{total} documents ({100 * affected / total:.1f}%) carry content Word 2010")
        print("  silently ignores. Every G0-G9 gate compares parsed models, and the parser")
        print("  does not model it either -- so those documents can pass every gate while")
        print("  having lost content. That is what the G10 preservation gate is for.")

    print(f"\nTOP {top} BY COMPLEXITY")
    print(f"  {'score':>6} {'struct':>6} {'rare':>6} {'fail':>6} {'layout':>6}  {'risk':<12} document")
    for row in summary["top"][:top]:
        rarity = f"{row['rarity_score']:.3f}" if row["rarity_score"] is not None else "   -  "
        print(f"  {row['complexity_score']:.3f} {row['structural_score']:6.3f} {rarity:>6} "
              f"{row['failure_score']:6.3f} {row['layout_score']:6.3f}  {row['risk_class']:<12} "
              f"{(row['source_ref'] or row['sha256'][:12])}")

    failures = summary["parser_failures"]
    print(f"\nDOCUMENTS OUR OWN PIPELINE COULD NOT HANDLE: {len(failures)}")
    for row in failures[:15]:
        reason = (row["model_error"] or row["extract_error"] or "").split("\n")[0][:90]
        print(f"  {(row['source_ref'] or row['sha256'][:12]):<44} {reason}")
    print()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, required=True, help="directory tree to scan for .docx")
    parser.add_argument("--db", type=Path, default=Path("C:/WordReplica-Automation/lab/lab.db"))
    parser.add_argument("--source", default="local", help="provenance label recorded per document")
    parser.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 2) - 1))
    parser.add_argument("--limit", type=int, default=None, help="scan at most N documents")
    parser.add_argument("--top", type=int, default=20)
    parser.add_argument("--no-model-parse", action="store_true",
                        help="skip the expensive DocxParser layer (roughly 2x faster)")
    parser.add_argument("--json", type=Path, default=None, help="also write the summary as JSON")
    args = parser.parse_args(argv)

    summary = scan(
        args.root, args.db,
        source=args.source, workers=args.workers,
        parse_model=not args.no_model_parse, limit=args.limit,
    )
    _print_report(summary, args.top)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(json.dumps(summary, indent=2, default=str), encoding="utf-8")
        print(f"summary written to {args.json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
