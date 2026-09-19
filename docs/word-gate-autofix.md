# Local Word gate repair

The Windows Word Release Gate now has a bounded repair supervisor on `automation-dev`. A failure on `main` or a release branch does not authorize agent edits. The original failed job remains failed, even when a repair succeeds.

## What runs automatically

1. Require a completed test failure; installation failures, interrupted runs and timeouts require review. Save the failed gate output under `C:\WordReplica-Automation\autofix\github-<run>-<attempt>`.
2. Check the runner account's Codex login, configured Golden sources/reports, Golden environment and existing fail-safe state.
3. Clone a candidate locally on `automation-dev`, remove its remote, and invoke `codex exec --sandbox workspace-write` without workflow publishing credentials. Sandbox or login failures stop; the supervisor never disables sandbox enforcement.
4. Ask Codex for a new unit regression only. The supervisor runs it and requires pytest exit 1 with actual JUnit failures, not collection errors or skipped tests.
5. Freeze that regression. Give Codex at most three fix attempts. The supervisor rejects changes to existing tests, workflows, gate scripts, release/security code and ownership helpers. Initial supported production directories are parser, renderers, domain and OPC; other diagnoses require review.
6. Independently run the regression, full pytest, the real Word release gate, and `RUN_GOLDEN_CODEX.ps1` for every configured Golden document. Require fresh reports, unchanged Golden source hashes and no previously passing gate regression. Existing `stop_required` decisions are honored.
7. Commit only a verified candidate. A separate credentialed workflow step checks that the branch still matches the original SHA and makes a non-force push.
8. Follow-up jobs verify the new SHA on Windows, Ubuntu and real Word, then attach `WordReplica autofix / Non-Word` and `WordReplica autofix / Real Word` statuses directly to that commit. This avoids relying on push events suppressed for `GITHUB_TOKEN` or dispatch registration on `main`.

No automatic main promotion, release signing or publication is performed. Golden evidence gathered on a dirty candidate is not a same-commit FULL PASS claim. Main promotion still needs the existing unchanged-source, same-commit double Golden pass.

## Local prerequisites

- Runner running interactively as the Windows user with Word and Codex installed and logged in. `codex login status` must succeed for that account.
- Existing `INSTALL_CODEX_AUTOMATION.cmd` setup, including `C:\WordReplica-Automation\.venv`.
- Every source listed in `codex_automation.json` present under `C:\WordReplica-Automation\golden`.
- Retained `golden_report.json` for each source, with matching SHA-256 and no active fail-safe. Missing evidence is a blocker, not permission to skip Golden.
- GitHub Actions token permitted to write repository contents and commit statuses. No new API key or personal token is required.

The workflow's early prerequisite step prints missing local items without turning an ordinary release test into a failure. The repair supervisor enforces them before asking the agent for edits.

## Stops and review

A commit may start only one repair chain, persisted under `state/word_gate_autofix/<base-sha>.json`; rerunning GitHub Actions cannot reset the three-attempt budget. Autofix commits carry `WordReplica-Autofix: true` and cannot start a fresh chain. A later failure on a repaired commit therefore requires review.

An uncertain diagnosis, changed regression, unsupported file change, full-suite failure, Golden regression or failed authentication stops with `status.json` and retained local evidence. A timeout also writes `state/word_gate_autofix_STOP.json`, blocking future chains pending process-ownership review. Never globally kill Word to resume. No automatic fail-safe reset is provided.

Candidate files and agent logs remain in the local per-run directory for review and are never uploaded automatically. They are not pruned automatically; after reviewing a stopped run, its retained candidate/log directory can be removed manually. Keep the small state records to preserve retry limits.

Codex requests use the existing ChatGPT login and its usage limits. A quota or authentication error stops the chain. This is bounded agent assistance, not a guarantee that arbitrary failures can be repaired unattended.
