# Golden Retention Must Ignore Non-Run Diagnostics

## Context

The field-format Golden run `20260816T050950Z_af539637` completed all 3,038 interactive events, wrote its report, moved the run into `diagnostics`, and then failed because the CLI could no longer read that report. Both the work run and final diagnostics run were gone.

The diagnostics root also contained pytest temporary directories named `pytest-full-*`. `prune_diagnostics` currently treats every directory as a Golden run. A directory without `golden_report.json` produces an empty report and is classified as a failed run. Since `pytest-*` sorts after timestamped Golden IDs, retention kept the pytest directories as the two newest failures and deleted the actual Golden run—including the report the caller was about to read.

## Goal

Retention must count and prune only Golden run directories that contain `golden_report.json`. Other diagnostics directories must remain untouched and must not consume success/failure retention slots.

## Chosen design

Build the candidate run list from immediate child directories containing a `golden_report.json` file. Parse each candidate report once and use that cached payload for pinned, success, and failure classification. Prune only those candidate run directories.

Directories without a report are outside Golden retention ownership. This is safer than matching a run-name regex and remains compatible if run ID formatting changes.

Malformed report files remain Golden candidates and classify as failures, preserving the current conservative cleanup behavior for damaged run artifacts.

## Regression test

Create two reported failed runs and one lexically newer `pytest-full-*` directory without a report. With `keep_failures=1`, assert that:

- the latest actual failed Golden run is kept;
- the older reported failure is pruned;
- the pytest directory still exists;
- the pytest directory is not returned as a retained Golden run.

The test must fail against the current implementation before production changes.

## Acceptance

- Focused retention test RED then GREEN.
- Existing retention tests remain green.
- Full local pytest suite passes.
- Repeated official Golden leaves its final report readable in diagnostics.
- No change to Golden source or Word process ownership behavior.

This infrastructure repair is part of the same field-format verification attempt; the field production change is not checkpointed until a valid real-Word report proves fidelity improvement.
