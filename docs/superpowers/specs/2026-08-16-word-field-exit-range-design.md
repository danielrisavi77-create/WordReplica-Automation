# Word Field Exit Range Design

## Problem

The retained Golden run `20260816T055626Z_42089d19` has one remaining G2 finding: paragraph 1369 contains an extra zero-length formatting run. Parsing source and output with the same `DocxParser` used by the audit shows the structural cause. The source has 26 `field_end` tokens distributed across 17 body paragraphs, while the output has all 26 tokens appended to the final body paragraph.

`InteractiveWordController._event_CreateField` duplicates `field.Result`, collapses that result to its end, and keeps it as the active insertion range. A Word field result ends immediately before the field-end marker. Collapsing the result therefore leaves the cursor inside the field. Every later insertion pushes the field-end marker forward, nesting later content and eventually accumulating all markers at the end of the document.

## Considered Approaches

1. **Move one Word character past the field-end marker after collapsing the result (selected).** This directly matches the observed Word range model and changes only the cursor position after field creation.
2. **Set an absolute range from `field.Result.End + 1`.** This depends on manually calculated document positions and is less robust across story ranges.
3. **Create or update all fields after the rest of the document.** This is a broad architectural change with much higher correctness and performance risk.

## Design

After the cached result is seeded and formatted, duplicate `field.Result`, collapse it to `WD_COLLAPSE_END`, and call Word Range `Move` with unit `WD_CHARACTER` and count `1`. The resulting collapsed range is immediately after the field-end marker and becomes `active_range` for subsequent events.

The move is a required correctness operation. If Word exposes no callable `Move`, the renderer raises instead of silently continuing with a known-corrupt insertion position. The call uses positional COM arguments for compatibility with older Word wrappers.

No field refresh is added. Cached results, active run formatting, header/footer behavior, and Word process ownership remain unchanged.

## Regression Test

A Word-like field result fake models a result range whose collapsed end remains before a structural field-end marker. The test creates a field and asserts that the renderer moves the collapsed range exactly one character and leaves `active_range` after the marker. Before the fix, the position remains before the marker and the test fails.

## Acceptance Criteria

- The field-exit regression test is observed RED before production changes and GREEN after the minimal fix.
- The existing cached-result formatting test remains green.
- The full pytest suite passes.
- The real Word Golden source hash remains `2cf2207f673f0ff1613176c9ac696fc1d935b4760331df132a1ef3ca238c56d8`.
- G2 changes from one finding to PASS, with field-end tokens no longer accumulated in paragraph 1369.
- Every previously passing gate remains passing; G7 remains PASS.
- Interactive runtime remains below 45 minutes.
- Commit and push occur only after measurable Golden improvement.
