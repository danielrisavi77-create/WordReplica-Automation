# Lekta + WordReplica Development Runner E2E Plan

**Goal:** Prove the real packaged WordReplica runner path locally on an unlicensed Word machine without weakening the production Authenticode release gate, then close the newly identified release-identity gaps.

## Task 1: Add a development-only unsigned runner seam

**Repository:** `C:\WordReplica-Automation\repo` on `automation-dev`.

**Files:**
- Create `BUILD_LEKTA_REPAIR_RUNNER_DEV.ps1`.
- Create `scripts/lekta_repair_runner_dev_entry.py`.
- Create a development-only loopback transport/entry module under `src/word_replica/runner/` if needed.
- Create focused unit tests for the development build, entrypoint and endpoint validation.

**Acceptance criteria:**
- Start with a regression test and record RED because the development executable path does not exist.
- The development build is a separate script and executable named `LektaRepairDev.exe`.
- It packages the same pinned public contract trust store and `fixer_ids.json` as the release runner.
- It runs the existing focused runner regression suite and a frozen `--self-test`.
- It never calls `Set-AuthenticodeSignature` and never emits `lekta-repair-runner-manifest.json`.
- It emits only `lekta-repair-runner-dev-manifest.json`, marked `developmentOnly: true`, and verifies the built artifact is not Authenticode `Valid`.
- The development entry parses the same tokenized executable filename as production and invokes the same `OneShotRunner` and Pure DOCX/Word-oracle package service.
- It accepts claim endpoint, status endpoint, output directory and state directory noninteractively only behind an explicit development-E2E flag.
- Development endpoints and signed claim download URLs must use HTTPS-form exact-loopback hosts `127.0.0.1` or `::1`; the development-only transport rewrites them internally to local HTTP with proxies disabled. Direct `http://` input, credentials, redirects and every non-loopback URL are rejected.
- GUI folder selection and self-deletion are disabled only for this explicit development-E2E invocation.
- Production `portable_entry.py`, production endpoint constants, production HTTPS transport and `BUILD_LEKTA_REPAIR_RUNNER.ps1` remain unchanged.
- Targeted tests pass, followed by the complete WordReplica pytest suite.

## Task 2: Add and run a black-box Lekta-to-EXE E2E harness

**Repositories:** Lekta `feature/repair-contract-v1-current-v2` and WordReplica `automation-dev`.

**Acceptance criteria:**
- Start with a Windows-only Lekta regression test/command that fails because no executable E2E exists.
- Generate a fresh ephemeral P-256 Repair Contract package outside Git from the immutable original `C:\Users\PC\Documents\Kalogjera - seminar Havel.docx`, the existing local reconstructed target under `C:\WordReplica-Automation\lab`, and the existing seven formatting requests.
- The Repair Contract and device lifecycle remain cryptographically signed; only Authenticode is omitted in development.
- Start loopback claim, source download, target download and signed status routes using real Lekta handlers/services.
- Build `LektaRepairDev.exe` with only the ephemeral public contract key, rename it exactly as the browser download code does, and launch it as a separate process through the explicit development-E2E seam.
- Observe exactly one claim, ordered `processing` then terminal `completed`, valid device signatures, output/report hashes and rejection of a second claim.
- Confirm exit 0, output DOCX exists, source SHA-256 is unchanged, G0-G10 pass, OpenAndRepair is false, fields-update equality passes and no unrelated Word process is terminated.
- Fail rather than skip or report green when the development artifact is absent.
- Keep package, reports, executable and generated key material outside Git under local diagnostics.
- Re-run the production Lekta preflight and prove it rejects the unsigned development artifact.

## Task 3: Harden production release identity and source attestation

**Repository:** Lekta feature branch, with matching WordReplica manifest fields when required.

**Acceptance criteria:**
- Start with RED tests demonstrating that a different valid Authenticode signer plus self-consistent adjacent manifest is currently accepted, and that dirty/unreviewed branch state is not rejected.
- Require an independently configured expected publisher thumbprint and expected Repair Contract key id; compare actual signature, manifest and expected values fail-closed.
- Record and validate WordReplica engine version and source commit in the runner manifest.
- Refuse production deployment from a dirty tree, the wrong branch, an unexpected commit or an unreviewed SHA; preflight remains read-only.
- Do not invent certificate identities. If no trusted publisher certificate exists, the gate remains correctly blocked.

## Task 4: Review and checkpoint

**Acceptance criteria:**
- Run fresh complete WordReplica and Lekta gates, required Word/oracle evidence and a final independent read-only review.
- Commit and push only WordReplica `automation-dev` and Lekta `feature/repair-contract-v1-current-v2`; never merge `main` or `master`.
- Do not deploy Lekta or Supabase.
- Production release remains blocked until a trusted Authenticode signing identity produces a `Valid` signed `LektaRepair.exe` and all release gates pass.
