# PR Audit Result Workflow

`PR_Audit_Result.xlsx` preserves the original Final PO columns and appends auditor result columns.

## Appended Columns

| Header | Meaning |
|---|---|
| `Source Row` | Original row number from `Final PO.xlsx`. |
| `Scope` | Scope inferred from matched create-pr-cd ECC output when possible. |
| `Audit Result` | Final classification. |
| `Reason Code` | Machine-readable reason code. |
| `Expected Item` | Expected ECC PBOM code and quantity. |
| `Expected Quantity` | Expected quantity for the submitted item code. |
| `Expected Subcontractor` | Subcontractor from generated ECC output. |
| `Normal Quantity` | Portion of submitted quantity accepted as normal. |
| `Duplicate Quantity` | Portion of submitted quantity treated as duplicate. |
| `Expected ECC Evidence` | Generated ECC file, worksheet, and row evidence for the site/DU. |
| `Matched ECC Evidence` | Generated ECC evidence for exact item-code matches. |
| `Explanation` | Human-readable explanation for the result. |

## Workflow

### 1. Preserve Final PO Source Row

The workbook reader stores the original row number from `Final PO.xlsx`.

### 2. Match Generated ECC Entitlement

The auditor matches Final PO rows to create-pr-cd ECC rows using:

- Final PO physical site code to ECC `Site ID*`
- Final PO DU/logical site code to ECC `Delivery Unit Code*`

### 3. Validate Subcontractor

The submitted subcontractor must match the generated ECC subcontractor for the site or item.

Mismatch result:

```text
Abnormal - Invalid PO
INVALID_SUBCON_CHANGED
```

### 4. Validate Item Code

The submitted item code must appear as a generated ECC `PBOM Code*` for the matched site or DU.

Mismatch result:

```text
Abnormal - Wrong PO
WRONG_LINE_ITEM_MAPPING
```

### 5. Resolve Quantity and Duplicates

Only otherwise-valid rows enter duplicate resolution.

Consumption order:

1. Dispatch Date
2. Request Number
3. Dispatch Order Number
4. PO Line Number

Quantity outcomes:

| Condition | Audit Result | Normal Quantity | Duplicate Quantity |
|---|---|---:|---:|
| Valid quantity is within remaining ECC entitlement | `Normal` | submitted quantity | 0 |
| Valid quantity exceeds entitlement | `Abnormal - Duplicate PO` | available portion | excess portion |
| Wrong or Invalid row | Wrong/Invalid classification | 0 | 0 |

## Code References

## Annotated ECC Copies

When `--annotate-ecc-output` is used, the auditor also copies source ECC workbooks into a timestamped output folder and appends these columns to each copied `details` sheet:

- Audit Status
- Audit Reason Codes
- Final PO Match Count
- Submitted Quantity
- Normal Quantity
- Duplicate Quantity
- Final PO Evidence
- Audit Explanation

Status values are `NORMAL`, `DUPLICATE`, `MIXED`, `INVALID`, `WRONG`, or `NOT_IN_FINAL_PO`. Original create-pr-cd output files are not changed.

## Code References

Implementation:

```text
skills/tx-pr-auditor/scripts/audit_final_po.py
```

Important functions:

| Function | Role |
|---|---|
| `workbook_reader` | Reads Final PO and generated ECC workbooks. |
| `expected_matcher` | Matches submitted rows to generated ECC entitlement. |
| `audit_engine` | Produces initial classification and reason code. |
| `duplicate_resolver` | Produces normal and duplicate quantities. |
| `report_writer` | Creates `PR_Audit_Result.xlsx`. |
| `annotated_ecc_writer` | Creates timestamped copied ECC workbooks with appended audit status. |
