# Exact Word Image Size Implementation Plan

**Goal:** Prevent Word from changing an explicitly reconstructed image width when both source dimensions and the aspect-ratio lock are known.

**Architecture:** Correct only the `SetImageSize` COM assignment order; retain all existing blueprint, asset, and audit behavior.

## Tasks

- [x] Confirm all G4 findings share one root cause: 23 width-only mismatches across 24 drawings.
- [x] Add a focused regression that records the Word shape assignment order.
- [x] Verify RED: the old implementation enables the lock before assigning both dimensions.
- [x] Temporarily unlock, assign width and height, and restore the requested lock in `finally`.
- [x] Verify targeted GREEN: 2 passed.
- [x] Verify the executor file: 59 passed.
- [x] Verify the full suite: 420 passed, 52 skipped.
- [x] Run real Word Golden `20260816T074359Z_63c5feae`.
- [x] Confirm G4 PASS with zero findings, prior PASS gates preserved, 73/73 pages, unchanged source hash, and `stop_required=false`.
- [x] Verify review RED for unspecified source lock and suppressed unlock/restoration failures.
- [x] Preserve the existing lock when unspecified and propagate terminal COM lock failures through the retry helpers.
- [x] Verify hardened targeted GREEN: 5 passed; executor file: 62 passed; full suite: 423 passed, 52 skipped.
- [x] Run hardened real Word Golden `20260816T075543Z_63c5feae`: G4 PASS 0, prior PASS gates preserved, 73/73 pages, unchanged source, 257.98-second interactive runtime.
- [x] Obtain final read-only re-review, then commit and push only to `automation-dev`.
