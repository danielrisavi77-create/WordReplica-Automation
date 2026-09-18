## Scope

- [ ] The change has a narrow, documented purpose.
- [ ] Generated environments, build outputs, DOCX/PDF diagnostics and local reports are not committed.
- [ ] No unrelated user Word process is terminated or modified.

## Verification

- [ ] Non-Word CI is green.
- [ ] Relevant focused regression tests were added or updated.
- [ ] For renderer/fidelity/release changes: the real Windows + Microsoft Word gate has been run and evidence is recorded.

## Release impact

- [ ] Version metadata remains consistent.
- [ ] Lekta Repair Contract compatibility is unchanged or explicitly documented.
- [ ] Any release artifact is tied to its source commit and SHA-256.
