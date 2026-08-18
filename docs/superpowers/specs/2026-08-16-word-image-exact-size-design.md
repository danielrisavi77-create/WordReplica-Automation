# Exact Word Image Size Design

## Problem

Golden run `20260816T073042Z_93328a24` had 23 G4 findings. All 24 image assets and every other drawing property matched; 23 drawings differed only in `width_emu`, by small amounts. The renderer set `LockAspectRatio` before assigning width and height. With the lock enabled, Word treated the later height assignment as a request to recalculate width from the embedded image ratio, replacing the exact source width.

## Design

When width and height are both explicit, capture the existing Word lock if the source lock is unspecified, temporarily set `LockAspectRatio` to false, assign the exact width and height, and restore the requested or captured lock state in `finally`. Preserve the existing behavior for one-dimensional sizing, where an enabled lock should continue to scale the other dimension. Required lock assignments use the existing rejected-COM-call retry path and propagate terminal failures instead of silently continuing with unknown geometry.

This changes only the order of four COM property assignments. It does not copy drawing XML, replace image assets, or weaken G4 comparison.

## Regression Test

A fake floating Word shape records property assignments. The test requires this exact sequence:

1. unlock aspect ratio;
2. set width;
3. set height;
4. restore the requested lock.

The test was RED with the old sequence (`lock true`, width, height) and GREEN after the fix. Review regressions additionally cover an unspecified source lock and failures during both temporary unlock and final restoration.

## Acceptance Evidence

- Targeted image tests: 5 passed.
- Interactive Word executor tests: 62 passed.
- Full suite: 423 passed, 52 skipped.
- Real Word Golden `20260816T075543Z_63c5feae`: G4 PASS with zero findings after review hardening.
- G0, G1, G2, G3, G5, and G7 remained PASS.
- Source/output remained 73/73 pages and the source hash remained unchanged.
- Interactive runtime was 257.98 seconds, below the 45-minute fail-safe.
