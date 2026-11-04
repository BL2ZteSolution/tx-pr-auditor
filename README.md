# tx-pr-auditor

Composite PR-audit skill for AI Worker Platform. Public inputs are Final PO and EPMS workbooks. Auditor owns orchestration: invoke the one shared create-pr-cd subprocess for mandatory TSS/TI ECC entitlement, then audit Final PO. The platform only resolves and approves dependencies; it does not interpret domain logic.

## Shared MW dependency - v1.2.0

There is no nested MW submodule. Platform deployments use the configured `createPrCdRoot` (normally sibling `skills/create-pr-cd`), approved once by the platform. Auditor declares compatibility with MW 4.1.0; future MW versions require a reviewed compatibility update, not a second checkout/pin.

The platform freezes dependency version/fingerprint in the job invocation and supplies trusted runtime approval in `TX_PR_AUDITOR_MW_DEPENDENCY`. Auditor rehashes the declared runtime files, checks version and available Git identity, and uses that same package for generation and PR-model lookup. Missing/incompatible/changed dependencies fail closed. Results record the actual dependency version, fingerprint and commit when available.

For standalone composite use, install compatible MW separately and configure that environment variable with a trusted approval JSON object: `skillId`, `version`, absolute `root`, `packageSha256`, `runtimeFiles`, and optional `commit`. Use the same normalized package-fingerprint algorithm as the platform. Do not put approval/root selection in user job parameters. There is no nested fallback or automatic remote-head update.

Run: `python src/main.py --input-manifest <workspace>/skill-input.json`.

Use `scripts/audit_final_po.py --pr-model <approved-mw>/Info/input/pr_model.xlsx` for focused audit development; this direct analysis CLI is not the composite execution/approval path.

Audit scope is resolved from column B of the approved PR Model before description heuristics. Existing EPMS-reference/PR-model entitlement reconstruction is retained, including historical Planning/Backoffice audit-only evidence. This migration does not newly enable Planning generation: mandatory generated scopes remain TSS/TI. Existing owner review of entitlement reconstruction remains pending.

See [migration validation and release boundary](docs/SHARED_MW_DEPENDENCY_1_2_0.md). Publish the Auditor and companion platform approval together; see platform documentation for observed activation status.
