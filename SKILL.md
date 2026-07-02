---
name: tx-pr-auditor
description: Validate Final PO.xlsx after create-pr-cd by auditing submitted TX PR/PO lines against EPMS.xlsx and pr_model.xlsx. Use when Codex needs to design, implement, run, or maintain the TX PR Auditor skill, including Final PO validation, workbook ingestion, canonical mapping, EPMS matching, expected entitlement generation, audit classification, duplicate resolution, and PR_Audit_Result.xlsx report generation.
---

# TX PR Auditor

Use this skill to validate `Final PO.xlsx` after `create-pr-cd` PR/ECC generation. The auditor is an independent business-audit skill that compares submitted Final PO rows against EPMS business facts and PR Model entitlement rules.

## Contract

`create-pr-cd` owns PR/ECC generation and the surrounding workflow. It may provide file paths and save outputs, but it must not contain TX audit business rules.

This skill owns workbook reading, data normalization, EPMS matching, business fact generation, PR Model resolution, expected item generation, audit execution, duplicate detection, and report generation.

Keep TX business rules out of the framework. Implement business rules only inside this skill.

## Inputs

Accept exactly three explicit file paths from the framework:

- `Final PO.xlsx`
- `EPMS.xlsx`
- `pr_model.xlsx`

Do not search the workspace for inputs. Read Final PO from worksheet `条目明细`. Read EPMS from worksheet `data`. Use the single PR Model workbook supplied by the user.

Current local fixture paths:

```text
.openclaw/skills/tx-pr-auditor/input/Final PO.xlsx
.openclaw/skills/tx-pr-auditor/input/EPMS.xlsx
.openclaw/skills/tx-pr-auditor/input/pr_model.xlsx
```

Read `references/current-inputs.md` before implementing field mapping or workbook readers.

## Pipeline

Process the workbooks as a batch pipeline:

1. Workbook Reader
2. Field Mapper
3. Canonical Builder
4. EPMS Matcher
5. Business Fact Builder
6. PR Model Resolver
7. Expected Item Generator
8. Audit Engine
9. Duplicate Resolver
10. Report Writer

The pipeline orchestrator is the only module that calls stage modules. Stage modules must not call each other directly.

## Script

Use the bundled CLI for the current baseline implementation:

```bash
python3 scripts/audit_final_po.py \
  --final-po "input/Final PO.xlsx" \
  --final-po-sheet "条目明细" \
  --epms "input/EPMS.xlsx" \
  --epms-sheet "data" \
  --pr-model "input/pr_model.xlsx" \
  --output "output/PR_Audit_Result.xlsx"
```

Install runtime dependencies from `requirements.txt` when needed:

```bash
python3 -m pip install -r requirements.txt
```

The script is intentionally conservative where the complete mapping matrix is not yet available. It must preserve evidence and use Invalid/Wrong classifications instead of guessing.

## Dataset Rules

Each stage produces one new immutable typed dataset. Do not mutate an earlier dataset and do not use shared mutable state between stages.

Use typed collections such as Final PO records, EPMS records, business facts, expected items, audit results, and final results. Avoid generic row maps once field mapping is complete.

Read `references/architecture.md` before implementing or changing pipeline modules.

## Audit Rules

The audit engine compares submitted records against expected entitlement. Read `references/business-logic.md` before implementing or changing audit, expected-item, duplicate, or reason-code behavior.

Follow this decision priority:

1. Subcontractor Validation
2. Invalid Business Domain
3. Wrong Mapping
4. Quantity Validation
5. Duplicate Evaluation

Generate expected entitlement before comparison. Expected item generation must not perform audit decisions.

## Duplicate Rules

Run duplicate resolution after audit validation. Process only valid claims for quantity consumption.

Consumption order:

1. Dispatch Date
2. Request Number
3. Dispatch Order Number
4. PO Line Number

Wrong PO and Invalid PO records must not consume expected quantity. Duplicate applies only to otherwise-valid claims that exceed expected quantity within the current `Final PO.xlsx` snapshot.

## Output

Generate `PR_Audit_Result.xlsx` with the submitted records and audit result columns.

Recommended result fields:

- Audit Result
- Reason Code
- Expected Item
- Expected Quantity
- Expected Subcontractor
- Normal Quantity
- Duplicate Quantity
- EPMS Evidence
- PR Model Evidence
- Explanation

## Constraints

- Keep the `create-pr-cd` framework as the workflow shell.
- Keep this skill independent from framework behavior.
- Keep source workbook paths explicit.
- Keep processing stage-by-stage and batch-based.
- Keep all stage outputs inspectable for debugging.
- Treat `TX_PR_Auditor_Business_Logic_Specification_v1.0` as the governing business-rule source.
