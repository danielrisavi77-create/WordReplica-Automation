# Word Image-Run Complex-Script Font Preservation

## Context

The clean GOLDEN #1 run at commit `d930b2c0` reduced G2 mismatches from 123 to 118. The first remaining G2 divergence is `24/runs/0/font_cs`, and the same pattern occurs only at paragraphs 24, 491, 494, and 497.

All four affected runs contain a drawing and no visible text. Their blueprints are correct: `ApplyRunProperties` supplies `font_cs: Cambria` immediately before `InsertImage`. The interactive renderer then calls Word's `InlineShapes.AddPicture`, but it never applies the active complex-script font to the returned image range. Word therefore leaves `inline.Range.Font.NameBi` at the style default, Times New Roman.

## Goal

Preserve the blueprint's complex-script font on Word image runs with one targeted property write, without changing image geometry, range movement, document content, or the fast table path.

## Non-goals

- Reapplying every run property to image ranges.
- Changing image insertion, floating-image conversion, or geometry sequencing.
- Post-processing DOCX XML.
- Addressing unrelated G2, G3, G4, G6, G8, or G9 mismatches in this iteration.

## Chosen design

Immediately after `InlineShapes.AddPicture` returns, and before an optional `ConvertToShape`, the `InsertImage` handler will inspect the active run properties. When `font_cs` is a non-empty value, it will set `inline.Range.Font.NameBi` to that value through the renderer's existing retry helper.

The write remains best-effort under the same exception-suppression policy already used for optional Word font properties. A failed optional `NameBi` assignment must not leave the active range in an invalid state or prevent image geometry from being applied.

Only `NameBi` is reapplied. This is the smallest change supported by the diagnostics and adds one COM property assignment per image. It avoids the extra calls and formatting side effects of replaying the full run-property set.

## Data flow

1. `ApplyRunProperties` records `font_cs` in the executor's active run properties.
2. `InsertImage` inserts the image with `InlineShapes.AddPicture`.
3. The handler writes the active `font_cs` to `inline.Range.Font.NameBi`.
4. The existing code duplicates and collapses the post-image range.
5. If requested, the existing floating conversion and geometry mutations continue unchanged.

Applying the property before `ConvertToShape` gives inline and floating images the same source run formatting while the image still exposes its inline range.

## Regression test

Add a focused unit test using the existing fake image document:

1. Set active run properties to `font_cs: Cambria`.
2. Insert an image whose fake range begins with `Font.NameBi = Times New Roman`.
3. Assert that the inserted image range ends with `Font.NameBi = Cambria`.

The test must fail before the production change and pass afterward. Existing image-geometry tests continue to protect mutation order and floating conversion behavior.

## Verification and acceptance

The iteration is accepted only if:

- the focused regression test is observed RED before the fix and GREEN after it;
- the full regression suite passes;
- a clean real-Word GOLDEN #1 run removes the four known drawing-only `font_cs` divergences, reducing G2 from 118 to at most 114;
- G0, G1, G5, and G7 remain PASS and no previously passing gate regresses;
- the Golden source hash remains unchanged;
- the runtime does not materially regress.

The change is eligible for an `automation-dev` checkpoint commit and push only after that measurable Golden improvement is confirmed. It does not by itself satisfy the overall two-consecutive-FULL-PASS completion gate.

## Safety

The implementation must not modify the Golden source DOCX, globally terminate Word, close unrelated Word documents, or run on `main`. All real-Word validation remains under the local Golden harness and its ownership rules.
