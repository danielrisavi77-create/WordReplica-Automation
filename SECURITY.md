# Security

## Supported development line

Security fixes should target the active reviewed development/release-candidate line first and then be promoted through the normal release gates.

## Reporting a vulnerability

Use GitHub's private security advisory mechanism for this repository when available. Do not attach private academic documents, signing private keys, access tokens, OAuth secrets or production service-role credentials to a public issue.

A useful report includes:

- affected commit/version;
- affected surface (desktop, runner, contract validation, installer, updater, Word ownership, etc.);
- minimal reproduction with synthetic files when possible;
- expected and actual safety behavior.

## Security boundaries

Word Replica treats these as security properties, not convenience behavior:

- source documents are read-only by default;
- signed repair packages fail closed;
- Repair Contract key IDs and fixer capabilities are allowlisted;
- runner contracts are bound to explicit file identities and engine compatibility;
- private signing keys are not packaged;
- unrelated Microsoft Word processes/documents must not be terminated or modified;
- process ownership uncertainty is a stop condition;
- release artifacts require hash/source traceability;
- a failed or unavailable fidelity check is never reported as a fabricated PASS.

## Local document privacy

The core Word Replica engine is local. The landing site is documentation/distribution only and does not need the user's DOCX.

External products integrating with Word Replica should exchange the minimum structured contract/status data needed for the workflow. Raw manuscript text should not be copied into shared backends merely to represent verification state.

## Signing

Development builds may be intentionally unsigned only in explicitly development-only flows.

Production executables must pass the repository's signing policy. Do not weaken the production Authenticode/Artifact Signing gate to make a development machine green.
