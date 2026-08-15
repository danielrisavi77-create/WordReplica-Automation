# WordReplica Word-Native Table Fast Path Design

**Date:** 2026-08-15
**Branch:** `automation-dev`
**Status:** User-approved design

## Objective

Reduce real-Microsoft-Word reconstruction time for table-heavy documents without weakening any fidelity, safety, audit, or restart requirement. The seminar must remain below five minutes with two consecutive G0-G9 FULL PASS runs. GOLDEN #1 must ultimately reach two consecutive G0-G9 FULL PASS runs in less than 45 minutes each.

Speed is accepted only when the reconstructed Word document remains content-, structure-, formatting-, pagination-, and render-equivalent under the existing G0-G9 gates. File byte size is not an acceptance criterion.

## Measured Evidence

The retained Golden trace from `20260815T115535Z_5ed0c311` contains:

- 10,722 completed events and 10,040.43 seconds of measured event time.
- 6,070.55 seconds (60.46%) inside tables.
- 5,078.09 seconds of `InsertText` time inside tables.
- 24 tables, 908 populated cells, and 928 table `InsertText` calls.
- The two slowest tables consume 1,510.40 and 1,282.93 seconds, or about 46.5 minutes together.
- The current renderer already combines ordinary content to approximately one text insertion per cell. Cell-level batching alone therefore cannot deliver the required improvement.

A static event-grammar scan found that all 24 Golden tables contain only table geometry, row/cell properties, one paragraph per cell, run formatting, and ordinary text. They contain no table-local images, fields, bookmarks, lists, tabs, line/page breaks, nested tables, or merge events. All 24 are candidates for the first fast-path implementation.

Even eliminating all table time would leave approximately 66 minutes of non-table Golden work. Table acceleration is therefore the largest first performance change, but a later document-wide text/layout optimization remains necessary for the 45-minute final target.

## Considered Approaches

### 1. Cell-level batching

Keep `Tables.Add`, enter each cell, and insert one combined text block per cell. This is safe but is already close to current behavior: the Golden trace has 928 text calls for 908 cells. It cannot materially change the scaling curve and is rejected as the primary optimization.

### 2. Word-native text-to-table transaction — selected

Build one delimited plain-text payload for an eligible rectangular table, insert it into the active Word range in one operation, and use Microsoft Word's native `ConvertToTable` operation to create the table. Apply table geometry, row/cell properties, paragraph properties, and run properties to exact Word ranges after conversion.

This reduces table text insertion from approximately one Word call per cell to one Word insertion and one conversion per table while still producing a real native Word table through Word COM.

### 3. Direct OOXML table injection

Write table XML directly into the DOCX package or copy source XML. This would be fast but would bypass the intended Word-native reconstruction path, complicate relationship handling, and provide weaker evidence about Word object-model equivalence. It is rejected for this optimization.

## Architecture

### Table batch plan

The blueprint compiler produces an immutable batch plan for every eligible table. The plan contains:

- source table ID, row count, and column count;
- table, column, row, and cell properties already resolved by the existing compiler;
- one text value per cell;
- one paragraph-format record per cell;
- ordered run spans with relative start/end offsets and effective run properties;
- the delimiter chosen for conversion, proven absent from all cell text;
- expected cell count, paragraph count, text projection, and run projection used for verification.

The batch plan is derived from the same parsed model and effective-formatting functions as the existing event stream. It must not introduce a second independent interpretation of styles or properties.

### Eligibility classifier

A table uses the fast path only when all of the following are known before Word mutation begins:

- rectangular row/column topology;
- no horizontal or vertical merge operation;
- exactly one paragraph per populated cell;
- text-only run content;
- no images, fields, bookmarks, notes, nested tables, list bindings, tabs, line breaks, or page breaks;
- a safe delimiter that does not occur in source text;
- all required paragraph and run offsets can be calculated deterministically.

Any table that fails one condition uses the unchanged legacy event path. Unsupported content is never silently dropped or simplified.

### Word transaction

For an eligible table, the renderer performs one atomic table transaction:

1. Declare a restart-safe table boundary, persist a checkpoint immediately before the transaction, and verify the ordinary pre-event Word state.
2. Insert one delimited text payload into the active range.
3. Convert that range to a native Word table with the exact planned dimensions.
4. Apply table properties, fixed column widths, row properties, and cell properties.
5. Resolve each cell's Word range and apply paragraph properties.
6. Apply each run's effective properties to its exact relative character range.
7. Re-establish the real paragraph after the table and the normal active-range state.
8. Verify the completed table projection before allowing subsequent document events.

The transaction remains visible when visible Word mode is requested. It appears as one table-level operation instead of hundreds of cell-level typing operations.

### Verification and observability

The fast path does not remove the Word state guard. It changes its granularity from every internal synthetic sub-event to the atomic table boundary and adds table-specific postconditions:

- expected Word table and cell counts;
- exact normalized cell text;
- expected paragraph count per cell;
- exact run text and effective formatting projection;
- table/row/cell geometry projection needed by G3;
- valid active range immediately after the table.

The retained trace records one table transaction with table ID, eligibility reason, row/cell/run counts, insertion time, conversion time, formatting time, verification time, total time, and fallback reason when applicable. Performance conclusions must be based on these retained timings rather than perceived UI speed.

### Failure and restart behavior

Eligibility is decided before mutation. There is no mid-transaction downgrade that would continue from a partially built table.

The execution service treats the batch as one restart unit. It must not record an internal batch phase as a completed resumable event; only the checkpoint immediately before the batch and the verified state after the complete batch are resumable boundaries.

If Word rejects or corrupts a batch transaction:

- the run stops with the exact table ID and failed phase;
- the source remains untouched;
- only the automation-owned Word document/process may be closed under the existing ownership proof;
- restart begins from the last proven safe checkpoint before that table;
- a deterministic compatibility rejection may route that table to the legacy path on the clean retry;
- unrelated Word processes and documents are never closed.

## Delivery Sequence

1. Add pure batch-plan and eligibility tests. Confirm RED before production code.
2. Add a deterministic table-heavy fixture with mixed normal/bold runs and exact geometry.
3. Add a real-Word microbenchmark for legacy and fast paths, including reopened G0-G3 projections.
4. Implement the smallest compiler representation for one eligible text-only table.
5. Implement Word-native insertion and `ConvertToTable`, then range-based formatting.
6. Verify targeted GREEN and the complete local pytest suite.
7. Run three benchmark pairs and accept only the median evidence.
8. Run two seminar reconstructions and full G0-G9 audits on the same commit.
9. Run one controlled Golden only after all preceding gates are green, then compare G0-G9, mismatch counts, page counts, source SHA-256, and runtime with the immediately previous retained run.
10. Continue with the first fidelity divergence or the next measured non-table performance bottleneck under `AGENTS.md`.

## Acceptance Gates

The table fast path is accepted only when all of these conditions hold:

- Targeted tests show observed RED before production implementation and GREEN afterward.
- Full local pytest passes.
- Reopened real-Word benchmark output has exact text, table/cell counts, paragraph/run boundaries, effective formatting, and geometry.
- Median table reconstruction time across three evaluable paired runs improves by at least 50%.
- No benchmark run loses G0-G3 fidelity.
- The seminar produces two consecutive G0-G9 FULL PASS reports on the same commit, remains at 17 pages, and reconstructs in less than five minutes.
- The Golden source SHA-256 remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- A controlled Golden loses no previously passing gate and materially improves table runtime or total runtime.
- G0-G9 tolerances are not loosened to make the optimization pass.

If the 50% table threshold is missed, the hypothesis is rejected without a production checkpoint. If a previously passing Golden gate regresses without an explained measurement artifact, execution stops under the existing fail-safe contract.

## Out of Scope

- Directly copying source table XML into the output package.
- Changing Golden source content or metadata.
- Weakening per-document safety, Word process ownership, or final G0-G9 auditing.
- Claiming the 45-minute Golden target from table work alone; retained evidence shows a separate non-table scaling problem remains.
