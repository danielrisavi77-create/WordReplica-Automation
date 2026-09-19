# Word gate autofix implementation plan

**Goal:** On an automation-dev Word gate failure, use the logged-in local Codex to prepare a regression and repair, independently verify it, then publish a checkpoint and run exact-commit verification jobs.

**Architecture:** Extend the existing self-hosted workflow after failure. A Python supervisor owns validation and commits. Codex runs sandboxed without GitHub credentials, prepares a new regression first and changes production code only after the supervisor observes RED. The original failed job stays failed. Only a verified new commit receives follow-up CI and Word jobs.

**Tech stack:** Python subprocess, existing Golden runner/locks, Codex CLI, Windows PowerShell, GitHub Actions.

**Spec:** Existing `docs/superpowers/specs/2026-08-12-wordreplica-codex-automation-design.md` and the user-approved chat design (bounded retries, local agent, no main promotion).

## Constraints
- automation-dev only; no automatic main/release promotion.
- Maximum three repair attempts per human commit; bot commits do not start another repair chain.
- No weakened existing tests, gate scripts, workflows, security or release code. Initially permit fixes only under parser/renderers/domain/opc and one newly added unit regression.
- Full local suite, Word release gate, and all configured Golden documents must be independently checked before commit. No Golden gate regression; unchanged source hashes; honor existing stop_required state.
- Missing authentication, Golden sources/reports, sandbox capability or evidence stops explicitly.
- Logs and candidate work remain locally outside Git; the agent receives no workflow token. Never globally terminate Word.

## Task 1: Supervisor and regression coverage
- [x] Add tests for allowed-path enforcement, Golden integrity/regression checks, observed RED, retry limits and bot-loop prevention.
- [x] Run `python -m pytest -q tests/unit/test_word_gate_autofix.py` and observe RED.
- [x] Implement supervisor with separate test and fix phases, inherited sandbox limits, native timeout fail-stop, and immutable test validation.
- [x] Repeat targeted tests to GREEN.

## Task 2: Workflow integration
- [x] Capture failed release output to a local per-run directory.
- [x] Run supervisor only after gate failure on automation-dev; use a clean local candidate clone without a credentialed remote.
- [x] Publish verified commit with a non-force push only if origin still matches base. Run dependent Non-Word and Word jobs against the exact new SHA and attach commit statuses; GITHUB_TOKEN pushes do not trigger push workflows and dispatch requires default-branch registration.
- [x] Retain original failure and emit a small Actions summary. Document preflight and stop semantics.

## Task 3: Verification and delivery
- [x] Run full Non-Word suite, compile checks and diff whitespace checks.
- [ ] Commit only implementation files to automation-dev and inspect triggered Windows/Ubuntu validation.
- [ ] Report separately code verification, real Windows execution and any local prerequisite blocker.
