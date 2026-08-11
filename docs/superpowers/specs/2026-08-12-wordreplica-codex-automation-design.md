# WordReplica Codex Automation Design

## Status
Approved architecture for automating WordReplica Golden-document development with a local Codex agent on the user's Windows machine, while using a private GitHub repository primarily for version control and backup rather than artifact storage.

## Goal
Remove the user's manual ZIP/update/harness loop. After one-time setup, Codex should iteratively analyze WordReplica Golden-run results, create regression tests, implement minimal fixes, rerun the real Microsoft Word reconstruction, and continue until the Golden document reaches a verified full-fidelity pass.

## Primary architecture

### Local Codex is the development loop
The fast loop runs on the user's Windows PC:

1. Codex reads the WordReplica repository and the latest Golden diagnostics.
2. Codex identifies the first concrete fidelity divergence.
3. Codex writes a failing regression test reproducing that root cause.
4. Codex verifies RED.
5. Codex implements the smallest production fix.
6. Codex verifies GREEN.
7. Codex runs the complete local regression suite.
8. If the suite passes, Codex launches a new real Microsoft Word Golden run.
9. Codex analyzes source-vs-output diagnostics immediately after the run finishes.
10. Codex repeats until the Golden gate is satisfied or a fail-safe condition stops the loop.

There is no hourly wait in the development loop. The test-analysis-fix-retest cycle happens locally and immediately.

### GitHub is not the runtime control plane
Repository: `danielrisavi77-create/WordReplica-Automation` (private).

GitHub is used for:
- source/version control;
- stable checkpoints;
- rollback history;
- optional final independent CI later;
- backup of code, tests, configuration, and small status reports.

GitHub is not used as the primary store for DOCX, PDF, PNG, or full event-trace artifacts during normal development.

## Repository strategy

### Branches
- `main`: stable versions only.
- `automation-dev`: Codex development branch.

Codex performs autonomous work only on `automation-dev`.

Promotion to `main` is allowed only after the Golden version passes the full defined gate twice on the same commit, proving repeatability.

### Commit policy
Codex should not commit every failed experiment.

Checkpoint commit is allowed when:
- a regression test and its fix are green;
- the full local regression suite is green; and
- the Golden result measurably improves or a previously failing Golden gate becomes green.

No checkpoint commit is required for a RED test or an abandoned hypothesis.

## Local directory layout
Recommended root:

```text
C:\WordReplica-Automation\
├── repo\                  # Git checkout
├── golden\                # local-only immutable Golden documents
│   └── Glavna verzija rektorova (grupno)(1).docx
├── work\                  # isolated per-run working directories
├── diagnostics\           # local source/output diagnostics
├── archive\               # retained selected prior runs
└── state\                 # machine-readable automation state
```

The Golden source document is never modified in place. Each Word run receives a copy inside a unique `work\run-<id>\` directory.

## Golden #1
Current Golden source:

`Glavna verzija rektorova (grupno)(1).docx`

The document is the primary fidelity benchmark because it exercises complex academic Word structures including long-form body text, styles, tables, images, headers/footers, fields, sections, page numbering, TOC/list structures, and pagination.

The current Golden audit model uses gates:
- G0 Content
- G1 Structure
- G2 Typography
- G3 Tables
- G4 Images
- G5 Page setup
- G6 Header/footer
- G7 Fields
- G8 Pagination
- G9 Visual fidelity

A `FULL PASS` requires all defined gates to pass.

## Codex operating contract
Codex receives a persistent project instruction equivalent to:

```text
GOAL:
Bring GOLDEN #1 to FULL FIDELITY PASS.

For every production fix:
1. Analyze the previous real Word run.
2. Identify the first concrete divergence/root cause.
3. Write a failing regression test.
4. Verify RED.
5. Implement the smallest safe fix.
6. Verify GREEN.
7. Run the full local regression suite.
8. If green, run the real Microsoft Word Golden reconstruction.
9. Audit source vs output.
10. Repeat only when evidence supports the next change.

Do not modify the Golden source document.
Do not claim PASS without measured evidence.
Do not skip the failing regression test.
Do not make unrelated fixes in the same iteration.
Do not globally terminate WINWORD.EXE.
Do not close the user's unrelated Word documents.
Do not push a regression to main.
```

## Microsoft Word isolation and safety
The user may have unrelated Word documents open while automation runs.

Requirements:
- create a dedicated Word COM instance for each automation run where feasible;
- track the process/window/COM instance that the automation owns;
- on timeout or crash, terminate only the automation-owned Word process if process ownership can be proven;
- never issue global commands such as `taskkill /IM WINWORD.EXE /F`;
- never alter or close unrelated user documents;
- enforce one active Golden Word run at a time.

## Diagnostics architecture
Diagnostics stay local by default.

Every run should create a machine-readable summary such as `golden_report.json` containing:
- run id;
- source hash;
- commit SHA;
- Word version/build;
- start/end/duration;
- reconstruction status;
- G0-G9 results;
- source/output page counts;
- first divergent gate;
- first divergent element/property where available;
- first/last completed reconstruction event;
- error category and traceback summary when applicable.

Heavy artifacts stay under `diagnostics\` locally:
- reconstructed DOCX;
- source/output PDFs;
- relevant page renders;
- full trace;
- OOXML diffs;
- QA intermediate files.

## GitHub storage policy
Normal development should consume approximately zero GitHub artifact storage.

### PASS
- no GitHub Actions artifact;
- code/test/status changes only when a checkpoint is warranted.

### Ordinary FAIL
- no automatic heavy upload;
- local `golden_report.json` and diagnostics are read directly by Codex.

### Optional escalation / later remote verification
If a future GitHub-based independent verifier requires artifacts:
- minimal diagnostic package target: <= 10 MB;
- visual/full escalation package hard limit: <= 50 MB;
- retention: 1 day;
- source Golden DOCX is never uploaded automatically;
- upload only failing pages or necessary excerpts, not all rendered pages.

## Local retention policy
Local diagnostics are automatically pruned to avoid disk growth.

Keep by default:
- latest successful run;
- latest two failed runs;
- any explicitly pinned baseline run.

Delete older unpinned run directories only after the current run has safely completed and its report is written.

## Fail-safe rules
The autonomous loop must stop and require review if any of the following occurs:
- three consecutive production iterations produce no improvement in Golden gates;
- a previously passing Golden gate regresses and the regression is not explained by the current intentional change;
- the same Word/COM failure repeats after its targeted retry/recovery policy is exhausted;
- the automation cannot prove ownership of a Word process it would need to terminate;
- the full local regression suite fails for an unrelated reason;
- diagnostics are insufficient to identify a concrete root cause;
- a proposed change would require destructive or system-wide actions.

## Golden promotion policy
`automation-dev` may be promoted to `main` only when:
1. all local unit/integration/regression tests pass;
2. Golden #1 reports G0-G9 PASS;
3. the same commit passes Golden #1 a second consecutive time;
4. the second run does not introduce new warnings or structural drift;
5. the source Golden hash is unchanged between the two runs.

After promotion, the stable commit is tagged/versioned according to the project's versioning convention.

## Future expansion
Only after Golden #1 is stable:
- add Golden #2 for complex table edge cases not covered by #1;
- add Golden #3 for floating/anchored image cases;
- add Golden #4 for advanced fields/notes/academic structures;
- run the original 20-document corpus as a regression suite;
- optionally add GitHub self-hosted Actions as an independent final verification gate, not as the primary development loop.

## Success criteria for this automation project
The setup phase is complete when:
- the private repo is cloned locally;
- Codex can operate on `automation-dev`;
- the Golden document is local-only and immutable;
- one command/task can execute the real Word Golden run end-to-end;
- Codex can read the generated report and local diagnostics without user upload;
- Codex can create RED/GREEN fixes and rerun automatically;
- failures trigger defined fail-safe behavior;
- stable versions are protected on `main`;
- normal development uploads no heavy artifacts to GitHub;
- the user no longer manually moves ZIP updates or uploads result ZIPs.
