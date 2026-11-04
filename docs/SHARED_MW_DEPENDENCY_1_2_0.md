# Shared MW dependency migration - Auditor v1.2.0

## Architecture

- Worktree `C:\deployment\ai-worker-platform-dev03`, platform branch `feature/auditor-shared-mw-dev03`, Auditor branch `feature/shared-mw-dev03`, based on platform `08efa9b`.
- Removed Auditor `.gitmodules` and nested `dependencies/create-pr-cd` gitlink only in this worktree. Production and dev01/dev02 are untouched.
- Platform resolves manifest-declared dependencies through existing configured approval roots. Auditor's dependency is shared MW 4.1.0 (`8fb4fd5bbda0ff1bbcf9b5cac5e339881f17e88e`), normalized package SHA-256 `f00d3a35b6321e301194d4380da6cf703fe961bf5d2d445af84530e28334a7c3`.
- No second MW package pin in Auditor: its manifest declares tested supported versions, while platform approval supplies actual fingerprint/files. Dependency cycle detection, incompatible version rejection and runtime revalidation are enforced.
- Freeze dependency version/fingerprint in the invocation at job creation; compare before subprocess start and again in Python. Old queued Auditor invocations without this metadata are blocked rather than silently upgraded. A new explicit rerun uses the current approved dependency.
- Pass trusted root/approval through TX_PR_AUDITOR_MW_DEPENDENCY, not arbitrary job parameters. Python rehashes files, validates entrypoint containment and optional Git commit, then uses the same root for child processes, PR model and upstream DU registry.
- Child job/result IDs, version, file containment and output hashes are verified. Actual dependency version/fingerprint/commit is retained in result entitlement metrics and catalog dependency metadata.
- EPMS XLSM support previously disagreed with MW's XLSX-only public adapter: now convert only a job-local macro-free copy for generation, leaving the uploaded source unchanged. XLSX is copied as before.
- Mandatory generated scopes remain TSS/TI. Existing audit-only PR-model/EPMS reconstruction remains unchanged; this is not a claim that Auditor was ECC-only or lacked historical Planning evidence.

## Validation

- Untouched Auditor baseline: 36 tests passed.
- Final Auditor: 44 tests passed, including actual shared TSS/TI subprocess + audit for synthetic XLSX/XLSM, immutable uploads, outputs/checksums, missing/unapproved/tampered/commit-mismatched dependencies, old frozen metadata rejection, cancellation and no Planning generation.
- Full backend npm suite passed with mock DB, isolated storage, disabled LLM and no live iEPMS. New platform acceptance creates real Auditor jobs using the shared MW, confirms Normal audit and output provenance, and blocks missing/changed dependency metadata before entitlement work.
- Frontend: 110 tests in 19 files passed; build and 13-route smoke passed. No frontend source changed.
- Focused shared MW regression: 53 passed / 12 subtests; one openpyxl extension warning. Includes TI new-link/reroute guidance and existing Planning adapter regressions. Not a full MW suite rerun.
- Before/after synthetic MW Swap comparison: old nested MW 4.0.0 and shared 4.1.0 both produce the same 2 TSS and 3 TI business lines (site/DU, contract/vendor, PBOM/SOW, unit/quantity). The same synthetic Final PO also retains Normal / NORMAL_FULL classification under both old and shared PR models. This fixture result does not prove all historical audits stay identical: newer governed TI rules intentionally differ.
- Registry SHA/parity source is unchanged between old and shared MW. No source rules were relaxed to make tests pass.
- Initial fixture-only failures (missing scope-specific raw export headers, test syntax, CRLF expected hash) were corrected and rerun; no production inputs edited.

## Release boundary

Auditor 1.2.0 publication and production restart are owner-authorized. Publish Auditor first, then the companion platform gitlink and matching approval. Revalidate the shared dependency and idle queue before restart; record activation only after live verification.

Roll back platform and Auditor pins together. Do not launch queued old-version envelopes against a new runtime. Standalone Auditor requires separately provisioned compatible MW and trusted approval. Deployment package need not contain Git metadata; package hashes remain mandatory and commit is recorded/checked when available.

Existing Final PO ingest caps and entitlement-boundary owner review remain unchanged and outside this path migration. No live PR submission or external business writes are included in this migration. Production activation is recorded separately after verification.

## Evidence

Retained under `C:\Development\aida-chrome\aida-temp-script\conversation-c-mumi1ss1-kgnh\dev03-validation`: baseline/final Auditor logs, backend/frontend logs, MW-focused log, original `auditor-old-mw-8d8880f.zip` and extracted old package, `mw-before-after-comparison.json` and `audit-before-after-comparison.json`. Synthetic platform acceptance artifacts remain under dev03 `command-center/runtime/auditor-shared-*`.

Auditor normalized approval fingerprint: `19e6eda24aaee4d4e72690689bfcc8dd170f58a1b63ce638c9f6623b5d8323f6`. Final backend rerun passed after all guards/negative checks; no nested dependency path remains in executable source or package approvals.
