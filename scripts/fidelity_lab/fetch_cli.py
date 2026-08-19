"""Fetch a corpus source onto disk, politely and resumably.

    python -m scripts.fidelity_lab.fetch_cli --source libreoffice
    python -m scripts.fidelity_lab.fetch_cli --source poi --limit 50

Both sources supported here are redistributable project test data (LibreOffice
is MPL-2.0/LGPL, Apache POI is Apache-2.0), so their bytes may be kept. A
source whose retention policy is DISCARD_BYTES is deliberately refused by this
command -- third-party documents get measured and dropped by the streaming
ingest, never written to disk in bulk.

Resumability is by existence: a file already present is skipped, so an
interrupted fetch costs only what it had not yet downloaded. The disk governor
is checked before every single write, because a guard that notices afterwards
has already caused the problem it exists to prevent.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
import time

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from word_replica.lab.budget import DiskBudget, DiskExhausted
from word_replica.lab.sources import (
    CircuitOpen,
    RetentionPolicy,
    apache_poi_source,
    libreoffice_source,
)

_SOURCES = {
    "libreoffice": libreoffice_source,
    "poi": apache_poi_source,
}


def fetch(source_name: str, dest_root: Path, *, limit: int | None, budget: DiskBudget) -> dict:
    source = _SOURCES[source_name]()
    if source.retention is not RetentionPolicy.KEEP_VERBATIM:
        raise SystemExit(
            f"{source_name} is {source.retention}; its bytes must not be written to disk in bulk"
        )

    dest_root = Path(dest_root) / source.name
    dest_root.mkdir(parents=True, exist_ok=True)
    budget.assert_can_start(dest_root)

    started = time.perf_counter()
    fetched = skipped = failed = 0
    written_bytes = 0
    errors: list[str] = []

    print(f"listing {source.name} ...", flush=True)
    refs = list(source.list_candidates(limit=limit))
    print(f"  {len(refs)} candidate documents", flush=True)

    for index, ref in enumerate(refs, start=1):
        target = dest_root / ref.source_ref
        if target.exists() and target.stat().st_size > 0:
            skipped += 1
            continue
        try:
            budget.assert_can_continue(dest_root)
            payload = source.fetch(ref)
        except (DiskExhausted, CircuitOpen) as exc:
            print(f"\nstopping after {fetched} documents: {exc}", flush=True)
            errors.append(str(exc))
            break
        except Exception as exc:  # a single bad document must not stop the batch
            failed += 1
            errors.append(f"{ref.source_ref}: {type(exc).__name__}: {exc}")
            continue

        if not budget.headroom_for(dest_root, len(payload)):
            print("\nstopping: writing this document would breach the disk floor", flush=True)
            break
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(payload)
        fetched += 1
        written_bytes += len(payload)
        if index % 100 == 0:
            print(f"  {index}/{len(refs)}  fetched={fetched} skipped={skipped} "
                  f"{written_bytes / 1024**2:.1f} MB", flush=True)

    elapsed = time.perf_counter() - started
    return {
        "source": source.name,
        "licence": source.licence,
        "candidates": len(refs),
        "fetched": fetched,
        "skipped": skipped,
        "failed": failed,
        "megabytes": round(written_bytes / 1024 ** 2, 1),
        "seconds": round(elapsed, 1),
        "dest": str(dest_root),
        "errors": errors[:20],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", choices=sorted(_SOURCES), required=True)
    parser.add_argument("--dest", type=Path, default=Path("C:/WordReplica-Automation/lab/docs"))
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--floor-gb", type=float, default=2.0,
                        help="refuse to start below this much free space")
    args = parser.parse_args(argv)

    budget = DiskBudget(floor_bytes=int(args.floor_gb * 1024 ** 3))
    try:
        summary = fetch(args.source, args.dest, limit=args.limit, budget=budget)
    except DiskExhausted as exc:
        print(f"ERROR: {exc}")
        return 2

    print(f"\n{summary['source']}: {summary['fetched']} fetched, {summary['skipped']} already present, "
          f"{summary['failed']} failed  |  {summary['megabytes']} MB in {summary['seconds']}s")
    print(f"  licence: {summary['licence']}")
    print(f"  into:    {summary['dest']}")
    for error in summary["errors"]:
        print(f"  ! {error}")
    return 0 if not summary["errors"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
