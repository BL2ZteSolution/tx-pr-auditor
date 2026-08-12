# TX PR Auditor Architecture

## Purpose

TX PR Auditor is a standalone composite Python skill. Its public entrypoint accepts Final PO and EPMS, generates mandatory TSS and TI ECC through a pinned `create-pr-cd` dependency, then validates Final PO against that generated entitlement.

## Product Boundary

`src/main.py` owns:

- Public contract validation for one Final PO and one EPMS workbook.
- Isolated `create-pr-cd` TSS and TI child runs.
- Dependency identity checks, progress forwarding, and cancellation propagation.
- Passing only generated ECC into the focused audit engine.
- Authoritative `result.json` assembly.

`scripts/audit_final_po.py` owns:

- Final PO and ECC workbook mapping.
- Canonical identity and quantity normalization.
- Entitlement, item, and subcontractor matching.
- Invalid, Wrong, Duplicate, and Normal classification.
- Deterministic quantity consumption and evidence output.

Only the generator dependency receives EPMS. The focused audit engine never reads EPMS or the PR model and never reconstructs entitlement.

## Runtime Flow

```text
Final PO + EPMS
  -> create-pr-cd 4.0.0 / TSS
  -> create-pr-cd 4.0.0 / TI
  -> generated ECC workbooks
  -> focused Final PO audit
  -> PR_Audit_Result.xlsx + summary + annotated ECC
```

Each generator scope runs in an isolated child workspace under `temp/create-pr-cd/<scope>`. The source inputs are copied; no uploaded workbook is modified.

## Focused Audit Pipeline

```text
Workbook Reader
  -> Field Mapper
  -> Canonical Builder
  -> Expected ECC Matcher
  -> Audit Engine
  -> Duplicate Resolver
  -> Report Writer
```

Only otherwise-valid claims consume expected quantity. Consumption is ordered by Dispatch Date, Request Number, Dispatch Order Number, and PO Line Number. Invalid and Wrong rows consume nothing.

## Safety

- Dependency skill ID and version must match the release declaration.
- Missing dependency packages and missing TSS/TI ECC fail closed.
- Generated and declared paths remain inside their workspaces.
- Cancellation reaches the active child process and the audit row loop.
- Original Final PO and EPMS files remain unchanged.
- The platform contains no entitlement or audit business rules.
