# TX PR Auditor Technical Design

## Document Control

| Field | Value |
|---|---|
| Skill | `tx-pr-auditor` |
| Status | Baseline implementation document |
| Runtime entrypoint | `scripts/audit_final_po.py` |
| Primary output | `output/PR_Audit_Result.xlsx` |
| Summary output | `output/PR_Audit_Result.summary.json` |
| Governing business reference | `references/business-logic.md` |

## Purpose

TX PR Auditor validates submitted TX PR/PO lines in `Final PO.xlsx` after `create-pr-cd` generation work. It compares submitted PO rows against EPMS operational facts and PR Model entitlement rules, then classifies each submitted row as:

- `Normal`
- `Abnormal - Invalid PO`
- `Abnormal - Wrong PO`
- `Abnormal - Duplicate PO`

The auditor does not generate ECC files. It is independent from `create-pr-cd/scripts/generate_tss_pr_ecc.py`.

## Package Layout

```text
tx-pr-auditor/
├── SKILL.md
├── requirements.txt
├── agents/
│   └── openai.yaml
├── docs/
│   └── TECHNICAL_DESIGN.md
├── input/
│   ├── Final PO.xlsx
│   ├── EPMS.xlsx
│   └── pr_model.xlsx
├── output/
│   ├── PR_Audit_Result.xlsx
│   └── PR_Audit_Result.summary.json
├── references/
│   ├── architecture.md
│   ├── business-logic.md
│   └── current-inputs.md
└── scripts/
    └── audit_final_po.py
```

## Runtime Dependencies

Runtime dependencies are listed in `requirements.txt`:

```text
openpyxl>=3.1
```

Install with:

```bash
python3 -m pip install -r requirements.txt
```

## CLI Contract

Run from `.openclaw/skills/tx-pr-auditor`:

```bash
python3 scripts/audit_final_po.py \
  --final-po "input/Final PO.xlsx" \
  --final-po-sheet "条目明细" \
  --epms "input/EPMS.xlsx" \
  --epms-sheet "data" \
  --pr-model "input/pr_model.xlsx" \
  --output "output/PR_Audit_Result.xlsx" \
  --summary-json "output/PR_Audit_Result.summary.json"
```

Default sheet names:

| Workbook | Default worksheet | Header row |
|---|---|---:|
| `Final PO.xlsx` | `条目明细` | 1 |
| `EPMS.xlsx` | `data` | 4 |
| `pr_model.xlsx` | `TX Line Item (After 21-Apr 26)` when present, otherwise first sheet | n/a |

If a required worksheet is missing, the script fails clearly and lists available sheet names.

## Input Contract

The caller must pass explicit paths. The auditor does not search directories for inputs.

Required inputs:

- Final PO workbook
- EPMS workbook
- PR Model workbook

Current local fixture paths:

```text
input/Final PO.xlsx
input/EPMS.xlsx
input/pr_model.xlsx
```

### Final PO Canonical Mapping

| Final PO header | Canonical field |
|---|---|
| `派工日期` | `dispatch_date` |
| `派工单号` | `dispatch_order_number` |
| `PO行号` | `po_line_number` |
| `需求单号` | `request_number` |
| `项目名称` | `project_name` |
| `项目编码` | `project_code` |
| `业务大类` | `business_domain` |
| `施工区域` | `region` |
| `采购区域` | `purchasing_area` |
| `分包商` | `submitted_subcontractor` |
| `逻辑站点名称` | `logical_site_name` |
| `逻辑站点编码` | `du` |
| `物理站点名称` | `physical_site_name` |
| `物理站点编码` | `site_code` |
| `外包代码` | `submitted_item_code` |
| `代码名称` | `submitted_item_description` |
| `量纲` | `submitted_unit` |
| `派工数量` | `submitted_quantity` |
| `结算数量` | `settlement_quantity` |
| `支付数量` | `paid_quantity` |
| `产品型号_备注` | `product_model_remark` |
| `派工单状态` | `dispatch_status` |
| `外包商编码` | `subcontractor_code` |

### EPMS Canonical Mapping

| EPMS header | Canonical field |
|---|---|
| `customer site code` | `site_code` |
| `customer site name` | `site_name` |
| `du code` | `du` |
| `region` | `epms_region` |
| `Province/State` | `province_state` |
| `Latitude (North Plus South Minus)` | `latitude` |
| `Longitude (East Plus West Minus)` | `longitude` |
| `TX Upgrade Scope` | `tx_upgrade_scope` |
| `BOQ Configuration` | `boq_configuration` |
| `Tx SOW` | `tx_sow` |
| `TX SOW Details` | `tx_sow_details` |
| `NE SOW Details` | `ne_sow_details` |
| `FE SOW Details` | `fe_sow_details` |
| `MW Config Antenna Size NE` | `antenna_size_ne` |
| `MW Config Antenna Size FE` | `antenna_size_fe` |
| `SubCon - TSS Team` | `expected_tss_subcontractor` |
| `Subcon PR - TSS` | `existing_tss_pr` |
| `SubCon - TI Team` | `expected_ti_subcontractor` |
| `Subcon PR - TI` | `existing_ti_pr` |
| `TX Cutover Date` | `tx_cutover_date` |
| `Subcon - Planning` | `expected_planning_subcontractor` |
| `Subcon PR - Planning` | `existing_planning_pr` |

EPMS contains repeated `actual end time` columns. The current baseline uses `TX Cutover Date` and the first non-empty repeated `actual end time` value as integration trigger evidence. Final operation trigger mapping should be confirmed before production hardening.

## Pipeline

The script implements a staged batch pipeline:

```text
Workbook Reader
Field Mapper
Canonical Builder
EPMS Matcher
Business Fact Builder
PR Model Resolver
Expected Item Generator
Audit Engine
Duplicate Resolver
Report Writer
```

### Workbook Reader

Functions:

- `read_table`
- `read_pr_model`
- `workbook_reader`

Responsibilities:

- Load explicit workbook paths.
- Require configured worksheet names for Final PO and EPMS.
- Read Final PO with header row 1.
- Read EPMS with header row 4.
- Read PR Model rows from `TX Line Item (After 21-Apr 26)` when available.
- Emit `RawDataset`.

### Field Mapper

Function:

- `field_mapper`

Responsibilities:

- Convert workbook headers into canonical field names.
- Preserve raw source rows.
- Emit typed `FinalPORecord` and `EPMSRecord` collections.

### Canonical Builder

Function:

- `canonical_builder`

Responsibilities:

- Normalize DU and submitted item code to text.
- Convert submitted and settlement quantities to numeric values.
- Build dispatch sort key.
- Derive `integration_end_date` candidate from EPMS trigger fields.

### EPMS Matcher

Function:

- `epms_matcher`

Responsibilities:

- Match Final PO rows to EPMS rows by DU.
- Infer audit scope from submitted item/domain.
- Attach expected subcontractor and EPMS business facts.

### Expected Item Generator

Functions:

- `expected_item_generator`
- `expected_items_for_record`
- `expected_items_from_pr_model`

Responsibilities:

- Generate expected entitlement before comparison.
- Apply fixed Planning and Operation rules.
- Extract mandatory TSS/TI items from PR Model by section and Tx SOW match.
- Apply create-pr-cd-style TI choose-group narrowing for AS/Expected Item generation.

Current fixed expected items:

| Scope | Expected code | Quantity |
|---|---|---:|
| Planning | `350001000403` | 1 |
| Operation | `350000592793` | 1 |

Current antenna item-code selection follows create-pr-cd behavior:

- Collect mandatory TI rows matching EPMS `Tx SOW`.
- Group choose-candidates by SOW, rule text, and choose category.
- Resolve antenna choose groups by selected NE/FE antenna size.
- Resolve region/material choose groups only when deterministic.
- Fail closed when create-pr-cd geography/route resolver behavior is required but not yet wired into the auditor.

Example:

```text
EPMS Tx SOW: MW New Link / Reroute
EPMS antenna size: 1.2m
AS Expected Item: 350001095410
AY PR Model Evidence: TX Line Item (After 21-Apr 26)!253
```

### Audit Engine

Function:

- `audit_engine`

Priority:

1. Missing EPMS business fact.
2. Subcontractor mismatch.
3. Unsupported or unknown domain.
4. Operation trigger missing.
5. No expected item generated.
6. Wrong item mapping.
7. Pending quantity validation.

The engine emits `PENDING_QUANTITY` for otherwise-valid rows so duplicate resolution can consume quantities globally in deterministic order.

### Duplicate Resolver

Function:

- `duplicate_resolver`

Responsibilities:

- Process only otherwise-valid claims.
- Sort by:
  1. Dispatch Date
  2. Request Number
  3. Dispatch Order Number
  4. PO Line Number
- Consume expected quantity by:

```text
DU + Scope + Submitted Item Code
```

Wrong PO and Invalid PO rows never consume quantity.

### Report Writer

Function:

- `report_writer`

Responsibilities:

- Preserve original Final PO columns.
- Append audit result columns.
- Write `PR_Audit_Result.xlsx`.
- Optionally write JSON summary.

Audit columns appended:

- `Source Row`
- `Scope`
- `Audit Result`
- `Reason Code`
- `Expected Item`
- `Expected Quantity`
- `Expected Subcontractor`
- `Normal Quantity`
- `Duplicate Quantity`
- `EPMS Evidence`
- `PR Model Evidence`
- `Explanation`

## Data Structures

The script uses frozen dataclasses to keep stage outputs immutable:

| Dataclass | Purpose |
|---|---|
| `RawDataset` | Raw workbook rows and metadata |
| `FinalPORecord` | Raw and canonical Final PO row |
| `EPMSRecord` | Raw and canonical EPMS row |
| `CanonicalDataset` | Canonical Final PO, EPMS, and PR Model rows |
| `BusinessRecord` | Final PO row plus EPMS match, scope, subcontractor, facts |
| `BusinessDataset` | Business-stage records |
| `ExpectedItem` | Expected item code, quantity, scope, evidence reason |
| `ExpectedRecord` | Business record plus expected entitlement |
| `AuditResult` | Per-row audit decision and quantities |
| `AuditDataset` | Audit result collection |

## Classification and Reason Codes

Current baseline outputs these classifications:

- `Normal`
- `Abnormal - Invalid PO`
- `Abnormal - Wrong PO`
- `Abnormal - Duplicate PO`

Implemented reason code families:

- `NORMAL_FULL`
- `NORMAL_PARTIAL`
- `INVALID_SUBCON_CHANGED`
- `INVALID_WRONG_DOMAIN`
- `INVALID_NO_OPERATION_TRIGGER`
- `INVALID_NO_EXPECTED_BUSINESS_FACT`
- `WRONG_ANTENNA_SIZE`
- `WRONG_LINE_ITEM_MAPPING`
- `DUPLICATE_PARTIAL_QUANTITY`
- `DUPLICATE_FULL_QUANTITY`

## Current Baseline Run

Latest recorded run:

```text
generated_at: 2026-07-01T19:54:50
total_rows: 35
Normal: 9
Abnormal - Invalid PO: 20
Abnormal - Wrong PO: 6
```

Reason code counts:

```text
INVALID_SUBCON_CHANGED: 1
NORMAL_FULL: 9
WRONG_LINE_ITEM_MAPPING: 6
INVALID_WRONG_DOMAIN: 5
INVALID_NO_EXPECTED_BUSINESS_FACT: 14
```

## Validation Commands

From repository root:

```bash
python3 -m py_compile .openclaw/skills/tx-pr-auditor/scripts/audit_final_po.py
python3 /Users/xmac/.codex/skills/.system/skill-creator/scripts/quick_validate.py .openclaw/skills/tx-pr-auditor
```

From `.openclaw/skills/tx-pr-auditor`:

```bash
python3 scripts/audit_final_po.py --help
python3 scripts/audit_final_po.py \
  --final-po "input/Final PO.xlsx" \
  --epms "input/EPMS.xlsx" \
  --pr-model "input/pr_model.xlsx" \
  --output "output/PR_Audit_Result.xlsx" \
  --summary-json "output/PR_Audit_Result.summary.json"
```

## Current Limitations

The baseline intentionally fails closed where the confirmed rule inputs are incomplete:

- Full TX SOW to Mapping Group to Expected Line-Item Set matrix is not yet encoded.
- PR Model effective-date worksheet selection is partially represented and should be hardened.
- Operation/integration trigger mapping from EPMS repeated columns requires final confirmation.
- TI legacy `0.5 + 0.5` antenna quantity behavior is documented but not fully implemented.
- Regional item logic is summarized in references but not fully encoded as a resolver.
- Acceptance tests for every classification and boundary condition still need to be added.

## Extension Plan

Recommended next steps:

1. Add a dedicated tests directory with workbook-free unit tests for scope inference, subcontractor normalization, expected item generation, and duplicate consumption.
2. Encode the full TX SOW mapping matrix as a structured reference file.
3. Implement PR Model worksheet effective-date resolver.
4. Confirm and encode EPMS operation trigger source columns.
5. Add regression fixtures for Normal, Invalid PO, Wrong PO, Full Duplicate, and Partial Duplicate.
6. Split `audit_final_po.py` into modules only after the rule surface stabilizes.
