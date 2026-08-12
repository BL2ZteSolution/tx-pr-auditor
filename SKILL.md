---
name: tx-pr-auditor
description: Run the complete TX PR audit from Final PO.xlsx and EPMS.xlsx. The standalone Python contract invokes its pinned create-pr-cd dependency for mandatory TSS and TI ECC entitlement, then validates Final PO rows, including subcontractor/item/quantity checks, duplicate detection, and PR_Audit_Result.xlsx generation.
---

# TX PR Auditor

## Platform Contract

For contract execution, initialize submodules recursively and use `python src/main.py --input-manifest <workspace>/input.json`. The public interface and limits are declared in `skill.json`; all generation and audit rules below remain skill-owned.

The public job accepts `Final PO.xlsx` and `EPMS.xlsx`. One Python entrypoint runs pinned `create-pr-cd` 4.0.0 for TSS and TI, then validates Final PO against those generated ECC workbooks.

## Contract

`create-pr-cd` owns EPMS/site-data interpretation, PR Model matching, contract lookup, and ECC generation.

`tx-pr-auditor` owns the composite sequence and downstream validation:

1. Run pinned `create-pr-cd` for mandatory TSS entitlement.
2. Run pinned `create-pr-cd` for mandatory TI entitlement.
3. Read submitted Final PO and generated ECC rows.
4. Compare, classify, and resolve duplicate quantity consumption.
5. Write `PR_Audit_Result.xlsx`, summary JSON, and optional annotated ECC evidence.

Only the dependency receives EPMS. The focused audit engine never reads EPMS or the PR model and must not reconstruct entitlement.

## Inputs

Accept explicit paths only:

- `Final PO.xlsx`
- One EPMS `.xlsx` or `.xlsm` workbook

Do not search the workspace for inputs.

Final PO defaults:

- Auto-detect worksheet `条目明细` with header row `1`
- Auto-detect worksheet `Sheet1` with header row `2`

Generated ECC defaults:

- Worksheet: `details`
- Header row: `1`

Read `references/current-inputs.md` before changing workbook mapping.

## Pipeline

Process the workbooks as a batch pipeline:

1. Workbook Reader
2. Field Mapper
3. Canonical Builder
4. Expected ECC Matcher
5. Audit Engine
6. Duplicate Resolver
7. Report Writer

The pipeline orchestrator is the only component that calls stage functions. Stage functions must not call each other directly.

## Script

Run from `skills/tx-pr-auditor`:

```bash
python scripts/audit_final_po.py \
  --final-po "input/Final PO.xlsx" \
  --expected-ecc "../create-pr-cd/output" \
  --output "output/PR_Audit_Result.xlsx" \
  --summary-json "output/PR_Audit_Result.summary.json"
```

Use repeated `--expected-ecc` arguments for multiple files or directories.

The default `config/du_registry.json` recognizes all nine unique create-pr-cd
DU identities. Use `--du-registry <path>` only to supply an explicitly managed
replacement registry.

Use `--filter-year <YYYY> --filter-month <1-12>` together to audit only Final PO rows whose Dispatch Date falls in that period.

Use `--final-po-sheet` and `--final-po-header-row` only to override format auto-detection.

Use `--final-po-max-rows <n>` only for bounded smoke tests or debugging large exports.

To generate copied ECC files with appended audit status columns, add:

```bash
--annotate-ecc-output \
--annotated-ecc-output-root "output"
```

Use `--annotated-ecc-timestamp <YYYYMMDD_HHMMSS>` only when deterministic folder names are needed for tests.

Install runtime dependencies from `requirements.txt` when needed:

```bash
python -m pip install -r requirements.txt
```

## Audit Rules

Use generated ECC rows as expected entitlement. Compare Final PO rows by:

- Site ID / physical site code
- DU / logical site code when available
- Registered create-pr-cd DU model when available in Final PO or ECC metadata
- PBOM/item code
- Subcontractor
- Quantity

Decision priority:

1. `Abnormal - Invalid PO`
2. `Abnormal - Wrong PO`
3. `Abnormal - Duplicate PO`
4. `Normal`

Invalid examples:

- No generated ECC entitlement exists for the submitted site or DU.
- Generated ECC entitlement has an unknown or conflicting DU identity.
- The same site/item entitlement exists in multiple DU models and Final PO
  does not identify the intended DU model.
- Submitted subcontractor differs from generated ECC subcontractor.

Wrong examples:

- Submitted site exists in generated ECC output, but the submitted PBOM/item code is not expected for that site.

Duplicate examples:

- Submitted item is valid, but submitted quantity exceeds generated ECC entitlement already consumed in this Final PO snapshot.

## Duplicate Rules

Run duplicate resolution after item/subcontractor validation. Process only otherwise-valid claims.

Consumption order:

1. Dispatch Date
2. Request Number
3. Dispatch Order Number
4. PO Line Number

Invalid and Wrong rows must not consume expected quantity.

## Output

Generate `PR_Audit_Result.xlsx` with the original Final PO columns plus audit columns:

- Source Row
- Scope
- DU Model
- DU Model ID
- DU Identity Key
- DU Profile ID
- DU Profile Status
- DU View ID
- Audit Result
- Reason Code
- Expected Item
- Expected Quantity
- Expected Subcontractor
- Normal Quantity
- Duplicate Quantity
- Expected ECC Evidence
- Matched ECC Evidence
- Explanation

When requested, also generate `PR_Audit_Result.summary.json`.

When `--annotate-ecc-output` is provided, also copy each source ECC workbook into:

```text
output/<YYYYMMDD_HHMMSS>/<original ECC filename>.xlsx
```

Append these columns to each copied ECC `details` sheet:

- Audit Status
- Audit Reason Codes
- Final PO Match Count
- Submitted Quantity
- Normal Quantity
- Duplicate Quantity
- Final PO Evidence
- Audit Explanation

Do not modify the original create-pr-cd output files. Label ECC rows with no exact Final PO item match as `NOT_IN_FINAL_PO`.

## Logging and Design Notes

Store revamp design notes, implementation decisions, and run logs under repository-level `./prompts`.

## Constraints

- Keep `create-pr-cd` as the PR/ECC generation owner.
- Keep this skill independent from EPMS and PR Model logic.
- Keep source workbook paths explicit.
- Keep processing stage-by-stage and batch-based.
- Preserve evidence for every audit result.
- Do not modify source workbooks.
