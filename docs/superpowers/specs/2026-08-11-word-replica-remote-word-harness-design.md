# Word Replica Remote Windows/Word Test Harness — Design

**Date:** 2026-08-11  
**Status:** Approved concept; written specification pending user review  
**Target:** Windows 10/11 + installed desktop Microsoft Word + Python 3.12+

## 1. Goal

Build a portable diagnostic harness that turns the user's Windows + Microsoft Word installation into a repeatable remote acceptance runner for Word Replica. The user drops 10–15 real `.docx` files into an input folder, runs one command, and receives one ZIP containing machine-readable and human-readable diagnostics for every document. The harness must continue after individual failures and must never modify source documents.

## 2. Why a separate harness

Three approaches were considered:

1. **Embed diagnostics in the desktop app.** Convenient, but couples experimental diagnostics to the release UI and makes failures harder to isolate.
2. **Wrap the existing pytest suite.** Fast to build, but pytest stops being an ideal artifact collector when Word COM crashes or a test process becomes unstable; it also produces fixture-centric rather than document-centric reports.
3. **Portable remote runner (selected).** A separate folder/tool that invokes the same production parser and rebuild services, but owns process isolation, timeouts, environment capture, artifact collection, result normalization, and final ZIP creation.

The selected design keeps release code and diagnostic orchestration separate while testing the real production engine.

## 3. User workflow

The package contains:

```text
WordReplica-Remote-Harness/
  START_HERE.cmd
  RUN_REMOTE_WORD_HARNESS.ps1
  realworld_input/
    PUT_DOCX_FILES_HERE.txt
  src/word_replica/
  scripts/remote_harness/
```

User workflow:

1. Extract the entire ZIP.
2. Copy 10–15 `.docx` files into `realworld_input/`.
3. Double-click `START_HERE.cmd`.
4. Harness checks prerequisites, then runs every document independently.
5. Harness creates `WordReplica-Remote-Results-<timestamp>.zip` next to the harness.
6. User uploads only that result ZIP back to ChatGPT.

No manual test selection is required.

## 4. Source protection

Every source document is treated read-only.

For each source file the harness records SHA-256 before and after the run. A source hash change is a harness-level `CRITICAL_FAIL` regardless of reconstruction outcome.

The harness never overwrites input files. Every reconstruction uses Word Replica's per-project output directories or a dedicated harness artifact directory.

## 5. Per-document test matrix

Each input document receives the following stages.

### 5.1 Static/package analysis

- filename and SHA-256
- file size
- OPC ZIP integrity
- `[Content_Types].xml` coverage
- relationship target existence
- parser success/failure
- canonical model summary
- capability/preflight classification
- counts: paragraphs, runs, tables, merged cells, images, drawings, sections, headers, footers, footnotes, endnotes, fields, bookmarks, comments/revisions, complex/unsupported parts
- blueprint compilation status and event counts

### 5.2 Instant baseline

Run the existing Instant Word reconstruction as a control. Capture status, warnings, reasons, output, QA report, save count, elapsed time, and source hash verification.

The Instant result is not allowed to hide an Interactive failure; it exists to distinguish parser/package problems from Interactive COM/executor problems.

### 5.3 Interactive Maximum attempt

Run Interactive Reconstruction with:

- Microsoft Word renderer
- visible Word
- maximum event speed
- object delay `0 ms`
- deterministic behavior
- preflight enabled
- live verification enabled where safe
- truthful save/audit behavior

If Maximum Fidelity correctly blocks an unsupported document during preflight, classify as `EXPECTED_BLOCK`, not test failure.

### 5.4 Interactive Standard fallback diagnostic

If Maximum is blocked only because of explicitly unsupported capabilities, optionally run Standard mode to learn whether supported portions execute. The report must retain explicit unsupported warnings and may never reclassify preserved/unsupported content as fully reconstructed.

### 5.5 Final QA

Where an output exists:

- reparse reconstructed DOCX
- L0 content comparison
- L1 structure comparison
- L2 formatting comparison
- L3 layout comparison where implemented
- Word→PDF L4 visual comparison where available
- package integrity and reopen checks

## 6. Process isolation and COM resilience

Each document/test mode runs in a fresh child Python process. This prevents one unstable Word COM session from poisoning the remaining corpus.

The parent runner:

- starts the child process;
- streams stdout/stderr to a per-run log;
- applies a configurable timeout;
- records exit code;
- kills only the child process tree on timeout;
- continues with the next document.

A Word COM failure such as `RPC_E_CALL_REJECTED`, `RPC server unavailable`, or a disconnected COM proxy is recorded verbatim with stage and traceback.

The harness does not silently retry arbitrary high-level document operations. Production-level COM retry/message filtering belongs in the renderer and is tested separately. The harness may retry only runner setup operations whose retry semantics cannot duplicate document edits.

## 7. Diagnostic event trace

Interactive runs emit an append-only JSONL trace containing, at minimum:

```json
{
  "event_index": 1024,
  "event_type": "InsertCharacter",
  "source_element_id": "r17",
  "story": "main",
  "section_index": 1,
  "table_depth": 0,
  "table_cell": null,
  "word_range_start": 183,
  "word_range_end": 183,
  "status": "before"
}
```

A matching `after` record is written after a successful atomic operation. On error the trace stores the failing event and exception type/message. This allows analysis to identify the exact first Word event that fails instead of only seeing a final `RunStatus.FAIL`.

No document text beyond what is necessary for diagnostics is required in the event trace by default. Character values may be included only when explicitly needed for reconstruction debugging; the result ZIP is intended to be user-controlled and uploaded voluntarily.

## 8. Environment capture

The result package records:

- Windows version/build
- architecture
- Python version
- Word version/build and bitness where discoverable through COM
- locale
- installed/default printer name if obtainable without elevated permissions
- project/harness version
- source-tree SHA-256 manifest
- relevant environment flags
- free disk space before/after

Do not collect unrelated personal machine data.

## 9. Output structure

```text
remote_results/<run_id>/
  summary.json
  summary.csv
  environment.json
  harness.log
  documents/
    001_<safe_name>/
      source_manifest.json
      analysis.json
      blueprint_summary.json
      instant/
        result.json
        stdout.log
        stderr.log
        output.docx              # when produced
        qa_report.html           # when produced
      interactive_maximum/
        result.json
        stdout.log
        stderr.log
        event_trace.jsonl
        output.docx              # when produced
        qa_report.html           # when produced
      interactive_standard/      # only when diagnostically useful
        result.json
        stdout.log
        stderr.log
        event_trace.jsonl
```

The final ZIP contains this tree exactly.

## 10. Normalized status taxonomy

Harness-level statuses:

- `PASS` — expected run completed and required QA passed.
- `WARN` — completed with explicit non-critical warnings.
- `EXPECTED_BLOCK` — Maximum Fidelity correctly rejected unsupported content before reconstruction.
- `ENGINE_FAIL` — Word Replica returned failure.
- `COM_FAIL` — Word/COM exception caused the run to fail.
- `TIMEOUT` — child exceeded configured duration.
- `PROCESS_CRASH` — abnormal child exit without a normalized result.
- `SOURCE_MUTATED` — source hash changed; critical harness failure.
- `HARNESS_FAIL` — diagnostic infrastructure itself failed.

Every status includes `stage`, `reason`, and artifact paths.

## 11. Aggregate summary

At completion print and write a compact table containing one row per document/mode:

```text
Document                     Instant   Interactive Max   Standard   First failure stage
seminar.docx                 PASS      COM_FAIL          -          BeginParagraph
thesis.docx                  PASS      EXPECTED_BLOCK    WARN       preflight/chart
images.docx                  PASS      PASS              -          -
```

Also group failures by normalized fingerprint so repeated root causes appear once, for example:

```text
RPC_E_CALL_REJECTED                 7 runs / 7 documents
INVALID_ACTIVE_RANGE_TYPE           3 runs / 2 documents
IMAGE_RELATIONSHIP_ERROR            1 run / 1 document
```

This grouping is critical: development should fix root causes, not documents individually.

## 12. Privacy and artifact policy

The harness runs locally. It does not upload anything automatically.

The user chooses whether to upload the final ZIP. The ZIP may contain copies of source-derived reconstructed documents and diagnostic metadata, so the tool must display a notice before collection and offer a `--logs-only` mode that excludes reconstructed DOCX/PDF artifacts if desired.

Source files themselves are not copied into the result ZIP by default. Their filenames, hashes, structural summaries, and generated outputs may be included. An explicit `--include-sources` option may be added later but is out of scope for v1.

## 13. Failure behavior

- One document failure never stops the corpus.
- Setup/prerequisite failure stops before running documents.
- Source mutation stops the affected document immediately and marks the overall run critical.
- Word process cleanup errors are recorded but may not overwrite an earlier, more specific failure.
- Missing output/QA artifacts are represented explicitly rather than causing artifact collection to crash.
- Final ZIP creation occurs even when every document fails, provided the harness process itself remains operational.

## 14. Prerequisite checks

Before running:

- Windows OS
- Python available or harness venv can be created
- sufficient free disk space
- desktop Microsoft Word COM activation succeeds
- input folder contains at least one `.docx`
- source tree/package integrity passes

If Word COM activation fails, stop immediately with a clear prerequisite error; do not run misleading non-Word Interactive tests.

## 15. Timeouts

Defaults:

- setup: 5 minutes
- static analysis: 2 minutes/document
- Instant reconstruction: 10 minutes/document
- Interactive reconstruction: 30 minutes/document in Maximum speed
- Word→PDF L4 stage: 10 minutes/document

Timeout values are configurable in `harness_config.json`.

## 16. No-deception constraint

The harness tests Word Replica's deterministic automated character-by-character reconstruction. It must not add random pauses, fake typing mistakes, simulated human corrections, fabricated edit history, or metadata intended to misrepresent automated work as manual authorship.

## 17. v1 success criteria

The remote harness v1 is complete when:

1. one command processes 10+ real DOCX files without manual intervention;
2. a failure in one child Word run does not prevent later documents from running;
3. each failure identifies the first failing stage/event when available;
4. source SHA-256 is unchanged for every document;
5. aggregate JSON and CSV summaries are produced;
6. a single result ZIP is produced on both all-pass and partial/all-fail corpus runs;
7. result ZIP contains enough evidence to distinguish parser, blueprint, executor, Word COM, QA, timeout, and harness failures;
8. no automatic upload occurs;
9. production Word Replica semantics are used rather than a mock renderer for Word stages.

## 18. Explicit non-goals for v1

- fixing the current Interactive renderer bugs inside the harness;
- cloud execution of Microsoft Word;
- automatic upload to ChatGPT or another service;
- benchmark/load testing many simultaneous Word instances;
- source document modification;
- human-like randomized typing simulation.

## 19. Relationship to current RC1 failures

The current Windows gate shows repeated `RPC_E_CALL_REJECTED` failures and active target/range type errors (`str` where a Word `Range` is required), plus a path-normalization unit-test mismatch. The harness is designed to make these and future root causes diagnosable across real documents without confusing repeated symptoms with independent document-specific bugs.

The harness does not treat these known failures as acceptable; it provides a better measurement loop for fixing them.

## 20. Versioning

The harness package must embed:

- harness version;
- Word Replica source version/hash manifest;
- schema version for JSON/JSONL artifacts.

Result ZIP consumers must be able to determine exactly which source snapshot generated the evidence.
