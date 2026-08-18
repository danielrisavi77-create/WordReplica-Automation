# Unexpected Header Shape Defaults Cleanup Design

## Problem

GOLDEN #1 run `20260816T093134Z_687e3788` passes G0 through G8 but fails G9. A controlled same-Word-instance probe compared the unchanged source, the accepted reconstructed output, and one-variable package copies. Restoring source run boundaries produced no pixel change. Replacing the whole `settings.xml` improved pages 4, 7, and 8, and the one-variable matrix proved that removing only output `w:hdrShapeDefaults` produces exactly the same improvement. Every other tested settings difference produced the baseline render.

The source has no `w:hdrShapeDefaults`; Word adds one to the reconstructed package. In the controlled matrix, removing it changed page 7 from changed-pixel ratio `0.037190049679804` and MAE `2.059900460899342` to ratio `0.025487075473015` and MAE `1.134638550011867`. Pages 29 and 70 were unchanged, so this is the first concrete G9 cause, not a complete G9 solution.

The seminar source and both accepted seminar outputs contain no `w:hdrShapeDefaults`. Its two accepted runs remain G0-G9 FULL PASS on commit `cabab84`, with 17 pages and reconstruction times of 55.43 and 55.00 seconds.

## Design

Add a deterministic package cleanup named `_remove_unexpected_header_shape_defaults(output_path, source_path) -> int` to `InteractiveRebuildService`.

- Read `word/settings.xml` from source and output after the owned Word reconstruction session has closed.
- If the source contains any direct `w:hdrShapeDefaults` element, leave the output unchanged in this iteration.
- If the source contains none, remove every direct output `w:hdrShapeDefaults` element.
- Preserve all unrelated settings and package parts byte-for-byte except for the rewritten output settings part.
- Return the number of removed elements and return zero without rewriting when no removal is required.
- Invoke the cleanup in `_finalize_qa` after the renderer closes and before L0-L4 auditing.
- Include its count when refreshing the final checkpoint hash and append `UNEXPECTED_HEADER_SHAPE_DEFAULTS_REMOVED` to the audit log when nonzero.

This change adds no Word COM calls, does not alter the source, does not copy the complete source settings part, and does not change fidelity tolerances.

## Verification

1. A regression fixture where source lacks the element and output contains it must observe RED before implementation, then verify removal while preserving another settings element.
2. A source fixture that contains `w:hdrShapeDefaults` must verify the output is untouched.
3. A finalization-level test must verify the cleanup is called after an owned renderer saves an output containing the unexpected element and records the audit event.
4. Run the complete pytest suite.
5. Run the real Word Golden and require G0-G8 to remain PASS, source SHA-256 unchanged, runtime below 45 minutes, and a measurable G9 improvement before committing.
6. Do not claim completion until GOLDEN #1 reaches two consecutive G0-G9 FULL PASS runs on one commit.
