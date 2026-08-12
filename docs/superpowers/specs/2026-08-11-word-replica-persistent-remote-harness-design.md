# Word Replica Persistent Remote Harness + Safe Updater Design

## Status
Approved design, formalized on 2026-08-11.

## Goal
Replace the current version-per-folder Remote Word Harness workflow with one persistent Windows harness directory that preserves the user's real-world DOCX corpus and historical results while allowing future harness versions to be applied as small, safe update packages.

## User-visible workflow
The user creates one permanent folder once, for example:

```text
C:\Users\PC\Downloads\WordReplica-Remote-Harness\
```

Its stable structure is:

```text
WordReplica-Remote-Harness\
  START_HERE.cmd
  UPDATE_HARNESS.cmd
  version.json
  harness_config.json
  realworld_input\
  results\
  current\
  backup\
  updates\
```

`realworld_input` is populated once with the user's 10–15+ DOCX files and is never overwritten or deleted by updates.

`START_HERE.cmd` always launches the harness from `current\` and writes each run to a new timestamped directory under `results\`.

Future ChatGPT-provided releases are small update ZIPs. The user places an update ZIP in `updates\` and runs `UPDATE_HARNESS.cmd`. The updater validates and installs the new code into `current\` without requiring a new harness folder or recopying input files.

## Protected persistent data
The updater MUST NOT modify, delete, move, rename, or replace:

- `realworld_input\**`
- `results\**`
- top-level `harness_config.json`
- any prior result ZIPs
- any user-created files outside `current\`, `backup\`, and `updates\`

These paths are persistent user data, not application code.

## Versioned application area
All replaceable harness code lives under:

```text
current\
```

This includes Python source, PowerShell runners, tests, package metadata, and internal manifests.

Top-level `START_HERE.cmd` is a stable launcher. It resolves the persistent root, then invokes `current\RUN_REMOTE_WORD_HARNESS.ps1` while passing the persistent `realworld_input` and `results` paths explicitly.

The current code must not assume that its own directory contains the input corpus or result history.

## Results layout
Every harness execution gets a unique run directory:

```text
results\
  2026-08-11_190500_v2\
  2026-08-11_193800_v3\
```

A run directory contains the complete per-document diagnostic artifacts for that execution and a result ZIP. Existing run directories are immutable from the updater's perspective.

The result ZIP name includes both timestamp and harness version, for example:

```text
WordReplica-Remote-Results-20260811-193800-v3.zip
```

This allows direct v2 → v3 → v4 comparisons on the same corpus.

## Version metadata
Top-level `version.json` records the installed harness version and update provenance. Minimum fields:

```json
{
  "schema_version": 1,
  "harness_version": "2.0.0",
  "installed_at": "2026-08-11T19:00:00+02:00",
  "package_sha256": "64-hex-character SHA-256 value",
  "previous_version": "1.0.0"
}
```

The harness also writes its version into every result summary.

## Update package format
A future update ZIP is not a full persistent harness. It contains only replaceable code plus an update manifest:

```text
WordReplica-Remote-Harness-Update-v3.zip
  update_manifest.json
  payload\
    src\
    scripts\
    pyproject.toml
    RUN_REMOTE_WORD_HARNESS.ps1
```

`update_manifest.json` contains:

- schema version
- target harness version
- minimum supported currently installed version
- SHA-256 for every payload file
- package identifier

The updater rejects malformed manifests, missing files, hash mismatches, incompatible version jumps, and archives containing paths outside `payload\`.

## Safe update algorithm
`UPDATE_HARNESS.cmd` invokes an ASCII-safe PowerShell updater.

The updater performs these steps in order:

1. Resolve the persistent harness root from the script location.
2. Locate exactly one selected/specified update ZIP in `updates\` or prompt for one if multiple are present.
3. Extract it into a temporary directory under the harness root.
4. Validate archive path safety and `update_manifest.json`.
5. Verify every payload SHA-256.
6. Read installed `version.json` and validate the version transition.
7. Run a non-Word update preflight against the extracted payload (Python import/compile plus updater-defined smoke tests).
8. Copy the existing `current\` tree to `backup\<timestamp>-<old-version>\`.
9. Install the payload into a new staging directory, never directly over the live `current\` tree.
10. Atomically swap staging into `current\` only after staging validation succeeds.
11. Write the new top-level `version.json` atomically.
12. Run a post-install non-Word smoke test from the new `current\` tree.
13. If post-install validation fails, restore the previous `current\` and `version.json` from the backup and report UPDATE FAIL.
14. On success, report the old and new versions and preserve the backup for manual rollback.

No successful update message may be printed before the new `current\` passes post-install validation.

## Rollback
A top-level `ROLLBACK_HARNESS.cmd` restores a selected most-recent backup of `current\` and its associated version metadata.

Rollback never changes `realworld_input` or `results`.

Rollback validates the restored code before reporting success.

## Initial migration from v2
The first persistent package is a one-time bootstrap package, not an update ZIP.

When the user extracts it, they may copy the existing 17 DOCX files into `realworld_input` once. After that, future versions use update ZIPs only.

Optionally, a migration command may import an existing v2 `realworld_input` directory into the persistent root, but it MUST copy rather than move and MUST refuse to overwrite an existing file with a different SHA-256.

## Python environment strategy
The persistent harness uses one reusable top-level virtual environment:

```text
.venv_harness\
```

It is not part of `current\` and is not deleted on every update.

After an update, the updater runs dependency synchronization only when the payload's dependency fingerprint differs from the installed one. `pip --no-cache-dir` is used to minimize disk use.

If dependency synchronization fails, the update is rolled back and the previous environment/current version remains usable where possible.

## Harness invocation contract
`current\RUN_REMOTE_WORD_HARNESS.ps1` accepts explicit persistent paths:

```text
-InputDir <persistent realworld_input>
-ResultsRoot <persistent results>
-HarnessVersion <version>
```

This separation makes the current code replaceable without risking user data.

## Failure behavior
A failed harness run does not modify installed code.

A failed update does not modify the active `current\` tree after rollback completes.

If an update is interrupted before the atomic swap, the old version remains active.

If interruption occurs after the swap but before post-install validation, the next updater launch detects an incomplete transaction marker and offers/executes recovery to the last known-good backup.

## Security and integrity
The updater:

- rejects `..`, absolute paths, drive-qualified paths, and symlink-like archive entries intended to escape the extraction root;
- verifies every payload file against the manifest;
- never executes scripts from an unvalidated payload;
- records installed package SHA-256/version metadata;
- does not download anything automatically from the internet;
- never uploads user documents or results automatically.

## Disk-space behavior
The updater checks free disk space before backup/extraction.

Temporary extraction/staging folders are deleted after success or rollback.

Backups are retained conservatively. A separate explicit cleanup command may later delete old backups, but the updater itself never silently removes the only rollback copy.

## Compatibility with existing harness behavior
The document-processing logic, per-document subprocess isolation, source hash checks, event traces, QA artifacts, result summaries, and logs-only mode remain unchanged except for receiving persistent input/results paths and harness version metadata.

This feature does not attempt to fix Word renderer failures. It only changes installation/update/result persistence workflow.

## Tests and release gates
The feature requires automated tests for at least:

1. protected paths remain byte-identical after an update;
2. old results remain present after multiple updates;
3. update manifest file hashes are verified;
4. malicious archive traversal is rejected;
5. failed preflight leaves current version untouched;
6. failed post-install validation performs rollback;
7. successful update changes `current` and `version.json` only after validation;
8. repeated updates do not require recopying DOCX input;
9. result directories include timestamp and harness version;
10. launcher passes explicit persistent input/results paths;
11. dependency environment is reused when dependency fingerprint is unchanged;
12. rollback restores previous code/version without touching inputs/results;
13. interrupted transaction recovery returns to a known-good state;
14. one full non-Word regression gate remains green.

The first Windows smoke test must demonstrate:

- the same persistent folder is used for two consecutive harness versions;
- the same real-world input corpus is used without recopying files;
- both versioned result directories remain available after the second run.

## Out of scope
- automatic internet updates;
- background update downloads;
- self-modifying installed Word Replica desktop EXE;
- deleting historical results automatically;
- changing or sanitizing the user's source DOCX files;
- solving current Word COM fidelity bugs as part of updater work.

## Transparency
The persistent harness and updater are development/test tooling. Version metadata and result reports must identify the harness version used. Updates must never fabricate test history or alter prior result records.
