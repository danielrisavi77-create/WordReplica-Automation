# Footer Font, Field Topology, and Stable Word Audit Design

## Problem

After G4 reached PASS, G6 had three findings. The first was the effective complex-script font of the PAGE footer field: source `Cambria`, output `Times New Roman`. The blueprint already carried `font_cs=Cambria`, but paragraph-style configuration applied only the Latin font name and omitted `Font.NameBi`.

Adding the missing style assignment removed the footer mismatch, but one real-Word run triggered the fail-safe because Word serialized all 26 body REF fields as `w:fldSimple` instead of the source's complex `fldChar`/`instrText` form. The fields remained visible and semantically real, but the structural and semantic gates correctly rejected the different topology.

Two more real-Word effects were then isolated. Word updated the relationship-free header's `STYLEREF` cached result from source `6. Rezultati` to `Prilozi`, and separate Word applications rendered the same unchanged source DOCX inconsistently during PDF audit. Finally, a tab inserted without reapplying active run properties formed a bare run and changed one PDF text partition.

## Design

1. Apply `font_cs` to paragraph styles through Word's `Font.NameBi` property, using the existing COM retry helper.
2. Parse actual fields with begin/separate/end stacks and hierarchical `fldSimple` records. Pair only equal field cardinality, parent topology, and normalized instruction identity.
3. Process deepest fields first. Rebuild supported fields from source structural runs, flags, and formatting while preserving output result nodes and tails. Normalize supported cross-paragraph or nested-instruction fields conservatively without blocking other fields.
4. Export source and output PDFs through one owned `DispatchEx` Word application so both sides share one font/layout context. Explicit test exporters retain precedence.
5. Restore only expected, relationship-free source header XML after the owned reconstruction Word session closes. Calculate relationship-part paths with `PurePosixPath`; skip any header with a `.rels` part in either package.
6. Route tabs through the same formatted insertion path as text so Word cannot leave a bare, reset-format run.
7. Leave the package untouched when complete field topology cannot be paired safely.

This does not weaken any gate or close unrelated Word sessions. It canonicalizes Word's alternate serialization only after the owned reconstruction session closes and makes the controlled PDF comparison use one renderer instance.

## Regression Evidence

- Paragraph-style complex-script font test observed RED, then GREEN.
- `fldSimple` topology tests observed RED, then GREEN for nested fields, split instructions, flags, run formatting, tails, unmatched topology, cross-paragraph fields, and nested fields inside instructions.
- The cleanup applied to a copy of fail-safe run `20260816T080916Z_687e3788` expanded 26 fields and restored both G1 and G7 to PASS.
- Three separate-app exports of the unchanged source produced bad/bad/good page text, while three exports in one Word application produced correct/correct/correct.
- Relationship-free header restoration made G6 PASS on a copy of real run `20260816T090627Z_687e3788`; direct and nested `.rels` skip tests are GREEN.
- Tab-formatting test observed RED, then GREEN; full executor file: 64 passed.
- Final full suite: 437 passed, 52 skipped.
- Read-only review: no Critical or Important findings.
- Real Word Golden `20260816T093134Z_687e3788`:
  - G0 through G8 PASS; only G9 remains;
  - source/output 73/73 pages and identical page text partitions;
  - source hash unchanged;
  - interactive runtime 254.40 seconds; audit 58.41 seconds;
  - score 9 and `stop_required=false`.
