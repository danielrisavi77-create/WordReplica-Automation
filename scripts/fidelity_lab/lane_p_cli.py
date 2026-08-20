"""Lane P -- reconstruct real documents and gate them, without Microsoft Word.

    python -m scripts.fidelity_lab.lane_p_cli --db <lab.db> --limit 200

For each document: parse it, rebuild it through the pure-docx renderer
(ReconstructionMode.INSTANT, RendererChoice.DOCX -- never the interactive path,
whose configured timeout is four hours per document), then run G0-G7 from the
existing model gates and G10 from the new preservation gate.

The number this exists to produce is the **G10-only** count: documents that
every model gate reports as a perfect reconstruction and that G10 says lost
content. That is the false-pass rate of the current gate set, measured rather
than argued.

Nothing here touches Word, so it does not contend with a Golden run for the
machine-wide lock.
"""
from __future__ import annotations

import argparse
import collections
from dataclasses import dataclass
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import time
import traceback

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from word_replica.config import RebuildOptions
from word_replica.domain.enums import (
    FidelityMode,
    MetadataMode,
    ReconstructionMode,
    RendererChoice,
    VisibilityMode,
)
from word_replica.qa.golden_audit import build_model_gates
from word_replica.lab.word_probe import RepairOutcome, lock_is_held, probe_repair
from word_replica.qa.preservation import build_preservation_gate

MODEL_GATES = tuple(f"G{index}" for index in range(8))


@dataclass(slots=True)
class Outcome:
    source_ref: str
    ok: bool
    model_pass: bool | None = None
    g10_pass: bool | None = None
    first_model_failure: str | None = None
    g10_divergence: str | None = None
    repair: str | None = None
    error: str | None = None

    @property
    def false_pass(self) -> bool:
        """Every model gate green, G10 red -- a silent loss."""
        return self.model_pass is True and self.g10_pass is False


def _rebuild_options() -> RebuildOptions:
    return RebuildOptions(
        renderer=RendererChoice.DOCX,
        visibility=VisibilityMode.BACKGROUND,
        fidelity=FidelityMode.FULL,
        metadata=MetadataMode.PRESERVE,
        reconstruction_mode=ReconstructionMode.INSTANT,
    )


def _evaluate(source: Path, service, options, *, with_word: bool = False) -> Outcome:
    from word_replica.parser.parser import DocxParser

    ref = source.name
    try:
        result = service.rebuild(source, options)
    except Exception as exc:
        return Outcome(ref, ok=False, error=f"rebuild raised {type(exc).__name__}: {exc}")
    if result.output_path is None or not Path(result.output_path).exists():
        reasons = "; ".join(result.reasons) or str(result.status)
        return Outcome(ref, ok=False, error=f"no output: {reasons}")

    output = Path(result.output_path)
    try:
        parser = DocxParser()
        gates = build_model_gates(parser.parse(source), parser.parse(output))
    except Exception as exc:
        return Outcome(ref, ok=False, error=f"model gates raised {type(exc).__name__}: {exc}")

    model_pass = all(gates[name].passed for name in MODEL_GATES if name in gates)
    first_failure = next((name for name in MODEL_GATES if name in gates and not gates[name].passed), None)

    try:
        # The rebuild ran with the default metadata policy, which drops a
        # source's custom document properties unless they are allow-listed.
        # That is intended behaviour, so it is declared rather than counted as
        # a silent loss.
        g10 = build_preservation_gate(
            source,
            output,
            custom_properties_dropped_by_policy=True,
            application_properties_rewritten_by_policy=True,
        )
    except Exception as exc:
        return Outcome(ref, ok=False, model_pass=model_pass, error=f"G10 raised {type(exc).__name__}: {exc}")

    divergence = None
    if not g10.passed and g10.first_divergence:
        divergence = str(g10.first_divergence.get("path", ""))[:120]

    # Word is the slow, serial part, so it is opt-in and runs last: a document
    # that already failed to rebuild has nothing to open.
    repair = probe_repair(source, output).value if with_word else None

    return Outcome(
        ref, ok=True, model_pass=model_pass, g10_pass=g10.passed,
        first_model_failure=first_failure, g10_divergence=divergence, repair=repair,
    )


_ORDERINGS = {
    # Hardest first finds where reconstruction breaks; simplest first is where a
    # *silent* loss can actually show, because G0-G7 have to pass for G10 to be
    # the only gate that fails. "spread" walks the whole complexity range.
    "hardest": "ORDER BY complexity_score DESC",
    "simplest": "ORDER BY complexity_score ASC",
    "spread": "ORDER BY substr(sha256, 1, 8)",
}


def _pick(db: Path, limit: int, source_filter: str | None, order: str = "spread") -> list[tuple[str, str]]:
    connection = sqlite3.connect(db)
    connection.row_factory = sqlite3.Row
    query = (
        "SELECT source, source_ref FROM document "
        "WHERE risk_class = 'VALID' AND holdout = 0 AND extractor_ok = 1 "
        "AND model_error IS NULL AND complexity_score IS NOT NULL "
    )
    params: list[object] = []
    if source_filter:
        query += "AND source = ? "
        params.append(source_filter)
    query += _ORDERINGS[order] + " LIMIT ?"
    params.append(limit)
    return [(row["source"], row["source_ref"]) for row in connection.execute(query, params)]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--db", type=Path, default=Path("C:/WordReplica-Automation/lab/lab.db"))
    parser.add_argument("--docs", type=Path, default=Path("C:/WordReplica-Automation/lab/docs"))
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--source", default=None, help="restrict to one corpus source")
    parser.add_argument("--order", choices=sorted(_ORDERINGS), default="spread")
    parser.add_argument("--json", type=Path, default=None)
    parser.add_argument("--word", action="store_true",
                        help="also ask real Word whether it repairs each rebuild (slow, serial)")
    args = parser.parse_args(argv)

    from word_replica.services.rebuild import RebuildService

    if args.word and lock_is_held():
        # AGENTS.md: one active Golden Word run at a time, machine-wide. The lab
        # always gives way; RUN_GOLDEN_CODEX.ps1 must never wait on it.
        raise SystemExit("a Golden Word run holds state/golden_run.lock; yielding")

    picks = _pick(args.db, args.limit, args.source, args.order)
    if not picks:
        raise SystemExit("no eligible documents; run scan_cli first")

    workspace = Path(tempfile.mkdtemp(prefix="lane-p-"))
    service = RebuildService(app_root=workspace)
    options = _rebuild_options()

    outcomes: list[Outcome] = []
    started = time.perf_counter()
    for index, (source_name, source_ref) in enumerate(picks, start=1):
        path = args.docs / source_name / source_ref
        if not path.exists():
            outcomes.append(Outcome(source_ref, ok=False, error="document not on disk"))
            continue
        try:
            outcomes.append(_evaluate(path, service, options, with_word=args.word))
        except Exception:
            outcomes.append(Outcome(source_ref, ok=False, error=traceback.format_exc(limit=1).strip()))
        if index % 25 == 0:
            print(f"  {index}/{len(picks)} ...", flush=True)

    elapsed = time.perf_counter() - started
    evaluated = [o for o in outcomes if o.ok]
    errored = [o for o in outcomes if not o.ok]
    false_passes = [o for o in evaluated if o.false_pass]
    both_pass = [o for o in evaluated if o.model_pass and o.g10_pass]
    model_fail = [o for o in evaluated if o.model_pass is False]

    print(f"\n{'=' * 78}")
    print(f"  LANE P -- {len(picks)} documents in {elapsed:.0f}s ({elapsed / max(1, len(picks)):.2f} s/doc)")
    print("=" * 78)
    print(f"  evaluated                      {len(evaluated):>5}")
    print(f"  could not be evaluated         {len(errored):>5}")
    print(f"  G0-G7 pass AND G10 pass        {len(both_pass):>5}")
    print(f"  G0-G7 fail (already visible)   {len(model_fail):>5}")
    print(f"  G0-G7 pass but G10 FAILS       {len(false_passes):>5}   <-- silent loss")
    if evaluated:
        rate = 100 * len(false_passes) / len(evaluated)
        print(f"\n  {rate:.1f}% of reconstructions that every existing gate calls perfect")
        print("  have actually lost package content.")

    if false_passes:
        print("\nWHAT WAS LOST (first divergence path)")
        for path, count in collections.Counter(o.g10_divergence for o in false_passes).most_common(12):
            print(f"  {count:>4}  {path}")
        print("\nEXAMPLES")
        for outcome in false_passes[:10]:
            print(f"  {outcome.source_ref:<52} {outcome.g10_divergence}")

    probed = [o for o in evaluated if o.repair]
    if probed:
        introduced = [o for o in probed if o.repair == RepairOutcome.REPAIR_INTRODUCED]
        counts = collections.Counter(o.repair for o in probed)
        print("\nDOES WORD REPAIR WHAT WE HAND BACK?")
        for name, count in counts.most_common():
            marker = "   <-- ours" if name == RepairOutcome.REPAIR_INTRODUCED else ""
            print(f"  {name:<20} {count:>4}{marker}")
        if introduced:
            print("\n  Word repaired these reconstructions although it opened their")
            print("  sources cleanly, so the damage is ours:")
            for outcome in introduced[:10]:
                print(f"    {outcome.source_ref}")

    if model_fail:
        print("\nFIRST FAILING MODEL GATE (documents G0-G7 already catches)")
        for gate, count in collections.Counter(o.first_model_failure for o in model_fail).most_common():
            print(f"  {gate}  {count}")

    if errored:
        print(f"\nCOULD NOT BE EVALUATED: {len(errored)}")
        for kind, count in collections.Counter(
            (o.error or "").split(":")[0][:70] for o in errored
        ).most_common(8):
            print(f"  {count:>4}  {kind}")

    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(
            json.dumps(
                {
                    "documents": len(picks),
                    "seconds": round(elapsed, 1),
                    "evaluated": len(evaluated),
                    "errored": len(errored),
                    "both_pass": len(both_pass),
                    "model_fail": len(model_fail),
                    "false_pass": len(false_passes),
                    "word_probed": len([o for o in evaluated if o.repair]),
                    "repair_introduced": len(
                        [o for o in evaluated if o.repair == RepairOutcome.REPAIR_INTRODUCED]
                    ),
                    "false_pass_examples": [
                        {"document": o.source_ref, "divergence": o.g10_divergence}
                        for o in false_passes[:50]
                    ],
                },
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"\nsummary written to {args.json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
