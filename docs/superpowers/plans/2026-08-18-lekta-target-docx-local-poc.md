# Lekta Target-DOCX Local POC Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Verify a Lekta-signed corrected target, reconstruct it visibly in real Microsoft Word, preserve the original and unrelated Word sessions, and emit a complete local fidelity report.

**Architecture:** A new repair-contract package validates canonical JSON, ES256-P1363 signature, time, engine, filenames and both artifact identities before Word starts. A repair-package service delegates only the signed target to the existing interactive reconstruction engine, copies the verified result to a safe user destination, audits target versus output through reusable G0-G9 gates and keeps a package binding for safe resume. Lekta fixer semantics never enter WordReplica.

**Tech Stack:** Python 3.12+, `cryptography>=50,<51`, `rfc8785>=0.1.4,<0.2`, pytest, pywin32, Microsoft Word COM, existing WordReplica interactive engine and Golden audit.

## Global Constraints

- Work only in `C:\WordReplica-Automation\repo` on `automation-dev`.
- Never develop directly on `main`.
- Read latest local Golden diagnostics before every production-fix iteration.
- Follow evidence, regression test, RED, smallest fix, GREEN, full pytest, real Word Golden, G0-G9 comparison.
- Never modify the Golden source DOCX or `C:\Users\PC\Documents\Kalogjera - seminar Havel.docx`.
- Never globally terminate `WINWORD.EXE` or close unrelated Word documents.
- Terminate only an exact owned PID with matching creation FILETIME evidence.
- Enforce one active Word reconstruction run at a time.
- Stop on unexplained gate regression, insufficient diagnostics or ownership uncertainty.
- Do not commit DOCX, PDF, PNG, JPG, ZIP, event traces or reports.
- Do not create an interim checkpoint commit. Commit only after full pytest is green, standard Golden #1 remains green, and the new Kalogjera POC gate becomes measurably green.
- Do not promote to `main` unless the standard Golden report says `automation_decision.promotion_ready=true` and all broader product gates are satisfied.

## Prerequisite

Complete Tasks 1 through 3 of Lekta plan `docs/superpowers/plans/2026-08-18-repair-contract-target-artifact.md`. Copy its regenerated language-neutral fixture and generated fixer allowlist into this repository only after their Lekta tests pass.

## File Map

- Modify `pyproject.toml`: pinned canonicalization and signature dependencies.
- Create `src/word_replica/repair_contract/{__init__,contract,signature}.py`: strict v1 parser and verification.
- Create `src/word_replica/repair_contract/fixer_ids.json`: generated Lekta allowlist.
- Create `src/word_replica/repair_contract/package.py`: explicit package paths, identities and path safety.
- Create `src/word_replica/repair_contract/binding.py`: durable source/target/contract resume binding.
- Create `src/word_replica/repair_contract/report.py`: machine-readable completion model.
- Create `src/word_replica/services/repair_package.py`: preflight, rebuild, copy, audit, recheck and resume orchestration.
- Create `src/word_replica/qa/golden_audit.py`: reusable G0-G9 implementation.
- Modify `scripts/codex_automation/audit.py`: compatibility re-export.
- Modify `src/word_replica/cli.py`: visible `repair-poc` and `repair-poc-resume` commands.
- Create `src/word_replica/services/output_handoff.py`: post-verification user handoff.
- Create focused unit tests under `tests/unit/test_repair_contract_*.py` and `test_repair_package_service.py`.
- Create `tests/fixtures/repair_contract_v1/{valid-contract.json,public-key.spki.b64url,fixer-ids.json}`.
- Create `scripts/run_lekta_poc.py` and root `RUN_LEKTA_POC.ps1`: explicit visible real-Word harness.

---

### Task 1: Verify the Lekta contract cross-language before parsing

**Files:**
- Modify: `pyproject.toml`
- Create: `src/word_replica/repair_contract/__init__.py`
- Create: `src/word_replica/repair_contract/contract.py`
- Create: `src/word_replica/repair_contract/signature.py`
- Create: `src/word_replica/repair_contract/fixer_ids.json`
- Create: `tests/unit/test_repair_contract_signature.py`
- Create: `tests/unit/test_repair_contract_schema.py`
- Create: `tests/fixtures/repair_contract_v1/valid-contract.json`
- Create: `tests/fixtures/repair_contract_v1/public-key.spki.b64url`
- Create: `tests/fixtures/repair_contract_v1/fixer-ids.json`

**Interfaces:**
- Consumes: raw JSON object and `Mapping[str, bytes]` of trusted key IDs to SPKI DER.
- Produces: `RepairContractV1`, `canonical_unsigned_bytes(raw)`, `verify_signed_contract(raw, trusted_keys)`.

- [ ] **Step 1: Copy only the regenerated public fixture**

Copy the three small text/JSON files from Lekta `tests/fixtures/repair-contract-v1/`. Do not copy any private key or DOCX.

- [ ] **Step 2: Write failing cross-language signature tests**

Test this interface:

```python
raw = json.loads(CONTRACT_PATH.read_text(encoding="utf-8"))
keys = {"fixture-2026-08-16": decode_spki(PUBLIC_KEY_PATH.read_text().strip())}
contract = verify_signed_contract(raw, keys)
assert contract.target_file_name == "Kalogjera - seminar Havel-popravljeno.docx"
```

Clone and mutate `targetSha256`, `targetSize`, `targetFileName`, request params and expiry. Each mutation must raise `RepairContractSignatureError` before returning a parsed contract. Unknown key ID, malformed base64url and a 63-byte signature must fail closed.

- [ ] **Step 3: Verify RED**

Run:

```powershell
python -m pytest tests/unit/test_repair_contract_signature.py tests/unit/test_repair_contract_schema.py -q
```

Expected: import failure because the package does not exist.

- [ ] **Step 4: Add maintained canonicalization and crypto dependencies**

Add:

```toml
"cryptography>=50,<51",
"rfc8785>=0.1.4,<0.2",
```

Install the editable test environment through the repository's existing bootstrap mechanism. Do not silently fall back to `json.dumps` for signed bytes.

- [ ] **Step 5: Implement signature-first verification**

`canonical_unsigned_bytes(raw)` removes only `contractSignature` and calls `rfc8785.dumps`. Decode the 64-byte P1363 signature into `r` and `s`, convert it to DER, and verify P-256/SHA-256:

```python
r = int.from_bytes(signature[:32], "big")
s = int.from_bytes(signature[32:], "big")
der_signature = encode_dss_signature(r, s)
public_key.verify(der_signature, payload, ec.ECDSA(hashes.SHA256()))
```

Reject unsupported algorithm, unsafe key ID, unknown trusted key and invalid signature encoding before strict contract parsing.

- [ ] **Step 6: Implement strict v1 parsing**

Use frozen dataclasses and an exact top-level key set containing source identity, target identity, requests, exceptions, policies and signature. Reject bool-as-int sizes, unsafe DOCX filenames, bad UUID/time/semver, unknown fixer IDs not present in `fixer_ids.json`, duplicate request IDs and incomplete required G0-G9 policy.

- [ ] **Step 7: Verify GREEN and fixture parity**

Run the focused tests. Expected: all PASS and the fixture target identity matches its signed data.

- [ ] **Step 8: Run full pytest and standard Golden #1**

Run:

```powershell
python -m pytest -q
Set-Location C:\WordReplica-Automation
.\RUN_GOLDEN_CODEX.ps1
```

Expected: full pytest passes and the latest standard Golden report remains FULL PASS. If any prior gate regresses, stop immediately and inspect local diagnostics.

---

### Task 2: Validate explicit package files and collision-safe output

**Files:**
- Create: `src/word_replica/repair_contract/package.py`
- Create: `tests/unit/test_repair_contract_package.py`
- Modify: `src/word_replica/domain/errors.py`

**Interfaces:**
- Consumes: `RepairPackageRequest(original_path, target_path, contract_path, public_key_path, output_dir)`.
- Produces: `ValidatedRepairPackage` with contract digest and verified source/target snapshots; `reserve_output_path`.

- [ ] **Step 1: Write failing package preflight tests**

Define expected dataclasses in the test:

```python
request = RepairPackageRequest(
    original_path=original,
    target_path=target,
    contract_path=contract_path,
    public_key_path=public_key_path,
    output_dir=tmp_path / "out",
)
validated = load_and_validate_package(request, now=NOW, engine_version="0.1.0")
assert validated.source_snapshot.sha256 == validated.contract.source_sha256
assert validated.target_snapshot.sha256 == validated.contract.target_sha256
```

Add cases for aliased source/target, wrong filename, size/hash mismatch, expired contract, engine mismatch, output directory as a file, reserved output name, traversal and collision suffixing. Inject a `word_started` spy and assert it remains false for every preflight error.

- [ ] **Step 2: Verify RED**

Run `python -m pytest tests/unit/test_repair_contract_package.py -q`. Expected: import failure.

- [ ] **Step 3: Implement package validation and path safety**

Resolve every explicit path once, require source and target to be distinct regular `.docx` files, require path basenames to equal signed filenames, verify signature before context, then verify both sizes and SHA-256 values.

Select output from signed `outputPolicy.suggestedFileName`. If it exists, use `name (2).docx`, then `(3)`, without overwriting. Reject an output path that aliases source or target after Windows case-insensitive normalization.

- [ ] **Step 4: Verify GREEN, full pytest and standard Golden**

Run focused tests, full pytest and `C:\WordReplica-Automation\RUN_GOLDEN_CODEX.ps1`. Require no regression before continuing.

---

### Task 3: Persist package binding and reject changed resume inputs

**Files:**
- Create: `src/word_replica/repair_contract/binding.py`
- Create: `tests/unit/test_repair_contract_binding.py`

**Interfaces:**
- Consumes: `ValidatedRepairPackage`, reconstruction project ID and output path.
- Produces: `RepairRunBinding` and `RepairRunBindingStore.create/load/validate`.

- [ ] **Step 1: Write failing binding tests**

Expected model:

```python
binding = RepairRunBinding.build(
    job_id=contract.job_id,
    contract_sha256=validated.contract_sha256,
    source_sha256=validated.source_snapshot.sha256,
    target_sha256=validated.target_snapshot.sha256,
    project_id="project-1",
    output_path=output,
    engine_version="0.1.0",
)
```

Round-trip JSON atomically. Mutating contract bytes, original, target, job ID, output path or engine version must make `validate` fail before Word resume is called.

- [ ] **Step 2: Verify RED**

Run `python -m pytest tests/unit/test_repair_contract_binding.py -q`. Expected: import failure.

- [ ] **Step 3: Implement atomic binding storage**

Store under the WordReplica app root at `repair_jobs/<job_id>/binding.json` using temporary-file then replace. Include schema version 1 and exact keys. Never store document text, token or private key.

- [ ] **Step 4: Verify GREEN, full pytest and standard Golden**

Run focused tests, full pytest and the standard real Word Golden. Stop on any unexplained regression.

---

### Task 4: Move G0-G9 audit into the package without changing behavior

**Files:**
- Create: `src/word_replica/qa/golden_audit.py`
- Modify: `scripts/codex_automation/audit.py`
- Create: `tests/unit/test_golden_audit_public_api.py`
- Modify: existing `tests/unit/test_codex_automation_*.py` only if import paths require it.

**Interfaces:**
- Consumes: target/output DOCX pair and existing render dependencies.
- Produces: unchanged `GateResult`, `build_golden_report`, `build_model_gates`, `build_visual_gate`, `audit_docx_pair` from the installable package.

- [ ] **Step 1: Write failing public API parity test**

```python
from word_replica.qa.golden_audit import audit_docx_pair as packaged
from scripts.codex_automation.audit import audit_docx_pair as legacy
assert legacy is packaged
```

Also call the packaged function with existing fakes and assert all G0-G9 results match the legacy expected report.

- [ ] **Step 2: Verify RED**

Run the new test plus all `test_codex_automation_audit*.py` tests. Expected: missing packaged module.

- [ ] **Step 3: Extract without redesign**

Move the current audit implementation unchanged into `word_replica.qa.golden_audit`. Replace the script module with explicit imports and `__all__` for the five public names above. Do not alter tolerances or projections.

- [ ] **Step 4: Verify GREEN, full pytest and standard Golden**

Require public API parity, full suite and standard Golden FULL PASS.

---

### Task 5: Orchestrate preflight, visible rebuild, audit and original recheck

**Files:**
- Create: `src/word_replica/repair_contract/report.py`
- Create: `src/word_replica/services/repair_package.py`
- Create: `tests/unit/test_repair_package_service.py`

**Interfaces:**
- Consumes: `RepairPackageRequest`, trusted key, `RebuildService`, packaged `audit_docx_pair`, binding store.
- Produces: `RepairCompletionReport`; `RepairPackageService.run(request)` and `.resume(request, job_id)`.

- [ ] **Step 1: Write failing orchestration tests with fakes**

Expected call boundary:

```python
service = RepairPackageService(
    rebuild_service=fake_rebuild,
    gate_auditor=fake_audit,
    binding_store=store,
    clock=lambda: NOW,
    engine_version="0.1.0",
)
report = service.run(request)
assert fake_rebuild.sources == [request.target_path.resolve()]
assert report.original_unchanged is True
assert report.full_pass is True
```

Assert invalid signature/hash never calls `rebuild`. Assert WARN/FAIL creates a retryable/failed report without copying a successful output. Assert output copy hash equals the reconstructed project output and never overwrites source or target. Assert audit compares target to local output, not original to local output.

- [ ] **Step 2: Verify RED**

Run `python -m pytest tests/unit/test_repair_package_service.py -q`. Expected: import failure.

- [ ] **Step 3: Implement completion report**

Use exact JSON fields for job ID, contract/source/target/output hashes and sizes, validation stages, project ID, blueprint fingerprint when available, timing phases, Word ownership evidence path, OpenAndRepair, visible text/Fields.Update, G0-G9, original unchanged, output path, status and reasons.

`full_pass` is true only when reconstruction PASS, original unchanged, required gates all true, OpenAndRepair false and text policies pass.

- [ ] **Step 4: Implement run and resume**

Call package validation first. Invoke `RebuildService.rebuild` with interactive reconstruction, visible Word, maximum fidelity, fast mode and table fast path. Copy the verified project output to the reserved destination, verify the copy, audit target versus destination, recheck the original snapshot last, write report atomically, then persist binding.

For resume, validate package and binding first, call `resume_interactive(project_id)`, and execute the same final copy/audit/recheck path.

- [ ] **Step 5: Verify GREEN, full pytest and standard Golden**

Run focused tests, full suite and standard real Word Golden. Compare latest retained report to the prior baseline and stop on regression.

---

### Task 6: Add visible CLI, safe stop/resume and user handoff

**Files:**
- Modify: `src/word_replica/cli.py`
- Create: `src/word_replica/services/output_handoff.py`
- Modify: `tests/unit/test_cli.py`
- Create: `tests/unit/test_output_handoff.py`

**Interfaces:**
- Consumes: explicit original, target, contract, public key and output directory arguments.
- Produces: `repair-poc`, `repair-poc-resume`, JSON report path and a post-verification `.docx` handoff.

- [ ] **Step 1: Write failing CLI and handoff tests**

Parse only explicit paths:

```python
args = build_parser().parse_args([
    "repair-poc", "--original", "original.docx", "--target", "target.docx",
    "--contract", "contract.json", "--public-key", "public.spki",
    "--output-dir", "out",
])
assert args.command == "repair-poc"
```

Test exit 0 only for FULL PASS, exit 1 for retryable stop and exit 2 for invalid/failure. Mock the handoff function and assert it runs only after a verified full pass.

- [ ] **Step 2: Verify RED**

Run selected CLI/handoff tests. Expected: parser rejects new commands.

- [ ] **Step 3: Implement commands and post-verification handoff**

`open_output_for_user(path)` uses Windows file association only after full verification, equivalent to `os.startfile(str(path))`. It never obtains or closes another Word instance. The real acceptance run must confirm the association opens Microsoft Word and the corrected document remains open.

Expose a development-only `--stop-after-event` option that injects `InteractiveRunControl.request_stop()` at a safe observer boundary. Require resume to use the same explicit package and job ID.

- [ ] **Step 4: Verify GREEN, full pytest and standard Golden**

Run focused tests, full suite and standard Golden. No commit yet.

---

### Task 7: Build the explicit real-Word POC harness

**Files:**
- Create: `scripts/run_lekta_poc.py`
- Create: `RUN_LEKTA_POC.ps1`
- Create: `tests/unit/test_lekta_poc_harness.py`

**Interfaces:**
- Consumes: only `C:\WordReplica-Automation\poc\kalogjera\package` and an explicit diagnostics directory.
- Produces: visible run, local output and `repair_poc_report.json` with no broad folder search.

- [ ] **Step 1: Write failing harness contract tests**

Assert the PowerShell entry point runs in the foreground, uses `automation-dev`, refuses dirty production files except the known design/plan files during development, requires all package paths, creates no DOCX under Git and never contains `taskkill /IM WINWORD.EXE` or process-name-wide termination.

- [ ] **Step 2: Verify RED**

Run `python -m pytest tests/unit/test_lekta_poc_harness.py -q`. Expected: missing scripts.

- [ ] **Step 3: Implement foreground harness**

The Python harness validates configuration, records commit SHA and source hash, calls the CLI synchronously, reads the completion report, writes diagnostics under `C:\WordReplica-Automation\diagnostics\repair-poc-<UTC>`, and retains explicit original/target/output paths without copying user data into Git.

The PowerShell wrapper sets no background process and invokes the repository Python environment visibly.

- [ ] **Step 4: Verify GREEN and all non-Word tests**

Run harness tests and full pytest. Expected: all PASS.

---

### Task 8: Run Kalogjera acceptance, interruption/resume and checkpoint commit

**Files:**
- External package: `C:\WordReplica-Automation\poc\kalogjera\package`
- External diagnostics: `C:\WordReplica-Automation\diagnostics\repair-poc-*`

**Interfaces:**
- Consumes: package produced by the Lekta plan.
- Produces: two consecutive FULL PASS reports on one commit candidate plus one verified interruption/resume report.

- [ ] **Step 1: Record immutable baseline evidence**

Record SHA-256 for the original, target and contract. Confirm branch is `automation-dev`, source document is not under Git, no other automation run is active, and any currently open user Word documents are listed only by ownership-safe diagnostics.

- [ ] **Step 2: Run a controlled interruption**

Run `RUN_LEKTA_POC.ps1` with the development stop option at a safe semantic event. Expected: retryable status, valid binding/checkpoint, original unchanged and no unrelated Word document closed.

- [ ] **Step 3: Resume the same package**

Run `repair-poc-resume` with the same job/package. Expected: completion continues after the saved event, does not restart completed semantic work, and produces FULL PASS.

- [ ] **Step 4: Verify an unrelated Word document survives**

With an unrelated Word document already open, run the POC. Compare its document identity and process creation FILETIME before and after. Expected: it remains open and unchanged. Do not close it as part of POC cleanup.

- [ ] **Step 5: Obtain two consecutive FULL PASS runs**

Run the foreground POC twice on the same uncommitted candidate tree without code changes. Both reports must show original unchanged, valid target signature/hash, OpenAndRepair false, text/Fields.Update equality, G0-G9 true and successful output handoff.

- [ ] **Step 6: Re-run full regression and standard Golden #1**

Run:

```powershell
Set-Location C:\WordReplica-Automation\repo
python -m pytest -q
Set-Location C:\WordReplica-Automation
.\RUN_GOLDEN_CODEX.ps1
```

Expected: full pytest passes; standard Golden #1 remains FULL PASS with unchanged source SHA; Kalogjera is the new measurable green gate.

- [ ] **Step 7: Create the only WordReplica checkpoint commit**

Stage only source, tests, tiny JSON/text fixtures, scripts, design and plan files. Confirm no DOCX/PDF/image/ZIP/report is staged. Commit on `automation-dev`:

```powershell
git commit -m "feat: reconstruct signed Lekta repair targets"
```

Push only after explicit user authorization and only to `origin/automation-dev`. Do not promote to `main`.

## Completion Gate

This plan is complete only when strict signed-package preflight happens before Word, resume binds source/target/contract, local output matches the canonical target under G0-G9, original and unrelated Word sessions remain untouched, interruption/resume works, two consecutive Kalogjera FULL PASS runs occur on the same commit candidate, full pytest is green and standard Golden #1 remains green.
