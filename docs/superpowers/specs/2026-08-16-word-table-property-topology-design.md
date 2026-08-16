# Word Table Property Topology Design

## Problem

Golden run `20260816T065317Z_253a8e71` moved the first divergent gate to G3 with 1,429 findings while preserving 73/73 pages. The findings are dominated by Word's OOXML serialization of effective table formatting:

- 788 generated cell-border property groups;
- 452 generated cell `shading_fill="auto"` values;
- 68 source table borders with `val="nil"` omitted by Word;
- remaining table width/grid and table-level property normalization.

The reconstructed visible table formatting is produced through Word, but Word materializes inherited table edges on cells, writes automatic shading defaults, and removes explicit nil/default nodes. The G3 contract requires canonical table geometry and cell-property topology, not merely a visually equivalent serialization.

`InteractiveRebuildService._restore_source_table_layout` already performs deterministic package cleanup after the owned Word session is closed. It currently handles only generated `tblGrid`, `tblLayout`, and `tcW` nodes, leaving the rest of the Word-normalized table properties in the final package.

## Considered Approaches

1. **Restore exact table property containers during existing final cleanup (selected).** Pair source/output tables after G1-equivalent reconstruction and restore only `tblPr`, `tblGrid`, `trPr`, and `tcPr` nodes. This preserves reconstructed content and objects while recovering canonical property topology.
2. **Normalize G3 to ignore missing/nil/auto and generated cell properties.** This would hide real internal-structure differences and weaken the fidelity gate.
3. **Reapply every property through COM.** Word would serialize the same effective properties again, add thousands of calls, and materially slow the 5:39 reconstruction.

## Design

Extend `_restore_source_table_layout` with one private local helper that synchronizes an optional property child from a source parent to its output counterpart:

- if both nodes are byte-equivalent after XML serialization, do nothing;
- if source omits the node, remove the output node;
- if source contains the node, deep-copy it into the output at the corresponding position, replacing any Word-normalized node;
- count each changed property container once.

For every paired body table, synchronize:

- table `tblPr` and `tblGrid`;
- every paired row `trPr`;
- every paired cell `tcPr`.

Before changing any node, compare the complete recursive table/row/cell topology. If source and output topology differ, leave the output package untouched instead of risking a shifted pairing. When a missing `trPr` must be inserted, place it after `tblPrEx` to preserve schema order.

Only property containers are restored. Table text, paragraphs, fields, bookmarks, drawings, relationships, and reconstructed table content are not copied. Source remains read-only. G1 already passes and establishes equal structural order for deterministic pairing.

## Regression Test

A synthetic source DOCX contains explicit table nil borders, grid widths, row properties, cell width, shading, and cell borders. Its output counterpart contains Word-materialized/missing/default variants plus distinct output text. The test calls `_restore_source_table_layout` and asserts all four property container types match the source exactly while output text remains unchanged.

Two review regressions also prove that a mismatched nested-table topology causes no package mutation and that an inserted `trPr` follows `tblPrEx`.

Before the fix, the existing cleanup leaves the materialized properties and the test is RED. After the fix, it is GREEN.

## Acceptance Criteria

- The topology regression test is observed RED and then GREEN.
- Existing table-layout cleanup tests remain green.
- Full pytest passes.
- Real Word Golden preserves G0, G1, G2, G5, and G7 PASS.
- G3 materially decreases, with the expected target of PASS and zero findings.
- Source hash remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- Output remains 73 pages and interactive runtime remains below 45 minutes.
- Checkpoint is committed and pushed only to `automation-dev` after Golden improvement.
