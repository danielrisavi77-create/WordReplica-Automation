# WordReplica Automation Contract

Primary goal: bring GOLDEN #1 to G0-G9 FULL FIDELITY PASS.

## Mandatory loop for every production fix
1. Read the latest local `golden_report.json` and the retained diagnostics under `C:\WordReplica-Automation\diagnostics`.
2. Identify the first concrete divergence and its root cause. Do not guess from downstream symptoms when an earlier divergence exists.
3. Write a regression test that fails specifically for that root cause.
4. Run the targeted test and record RED.
5. Make the smallest production change that addresses the demonstrated root cause.
6. Run the targeted test and record GREEN.
7. Run the full local pytest suite.
8. Only if the full suite is green, run `RUN_GOLDEN_CODEX.ps1` from the repository root.
9. Compare the new G0-G9 report with the previous retained run.
10. Continue immediately when the next change is supported by evidence; otherwise stop under the fail-safe rules below.

## Safety
- Never modify the Golden source DOCX.
- Never globally terminate `WINWORD.EXE` and never use `taskkill /IM WINWORD.EXE /F`.
- Never close or modify unrelated user Word documents.
- Only terminate a Word PID when `owned_word.json` proves the exact PID, current automation child owner, and recorded Windows process-creation FILETIME, and the live process is still `WINWORD.EXE` with the same creation FILETIME.
- Enforce one active Golden Word run at a time.
- Never work directly on `main`; autonomous work happens only on `automation-dev`.
- Stop after three consecutive production iterations with no Golden gate improvement.
- Stop immediately on an unexplained regression of a previously passing Golden gate.
- Stop when diagnostics do not identify a concrete root cause or Word process ownership cannot be proven.
- Do not claim FULL PASS until the same commit passes G0-G9 twice consecutively with unchanged Golden source SHA-256.

## TDD
Every production behavior change requires: failing regression test -> observed RED -> minimal fix -> observed GREEN -> full regression suite -> real Word Golden run.
Do not commit failed hypotheses.

## Git
- `main` is stable only.
- Autonomous branch: `automation-dev`.
- Create a checkpoint commit only after tests are green and the Golden evidence measurably improves or a failing gate becomes green.
- Promotion to `main` is allowed only when `golden_report.json` says `automation_decision.promotion_ready=true`.

## Data and storage
- Golden and heavy diagnostics stay outside Git under `C:\WordReplica-Automation`.
- Never add DOCX, PDF, PNG, JPG, ZIP, `event_trace.jsonl`, or `golden_report.json` to Git.
- GitHub is version control/backup, not the normal diagnostic artifact store.
- Local retention keeps the latest successful run, latest two failed runs, and any explicitly pinned baseline.

## Golden #1 gates
- G0 Content
- G1 Structure
- G2 Typography and paragraph/run formatting
- G3 Table geometry and cell properties
- G4 Images and drawing geometry
- G5 Page setup and sections
- G6 Headers and footers
- G7 Fields, bookmarks, footnotes/endnotes
- G8 Exact pagination and page-text partition
- G9 Visual fidelity at 144 DPI with changed-pixel ratio <= 0.001 and MAE <= 0.25

## How to run
From the repository root on the configured Windows machine:

```powershell
.\RUN_GOLDEN_CODEX.ps1
```

Then read the printed report path and the JSON itself. Do not ask the user to upload ZIPs or manually move update packages.
