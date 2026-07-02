# PR Audit Result AO-AZ Workflow

This document explains the audit columns appended to `PR_Audit_Result.xlsx`.

The workbook structure is:

```text
Final PO original columns A-AN
+
Auditor result columns AO-AZ
```

## Column Meaning

| Column | Header | Meaning |
|---|---|---|
| AO | `Source Row` | Original row number from `Final PO.xlsx` sheet `条目明细`. |
| AP | `Scope` | Auditor-inferred scope: `TSS`, `TI`, `PLANNING`, `OPERATION`, or `UNKNOWN`. |
| AQ | `Audit Result` | Final classification: `Normal`, `Abnormal - Invalid PO`, `Abnormal - Wrong PO`, or `Abnormal - Duplicate PO`. |
| AR | `Reason Code` | Machine-readable reason, such as `NORMAL_FULL`, `INVALID_SUBCON_CHANGED`, or `WRONG_LINE_ITEM_MAPPING`. |
| AS | `Expected Item` | Expected item code or code set generated from EPMS and PR Model using create-pr-cd-style item selection. |
| AT | `Expected Quantity` | Expected quantity for the submitted item code when it matches expected entitlement. Often `0` for Wrong or Invalid rows. |
| AU | `Expected Subcontractor` | Expected subcontractor from EPMS based on scope. |
| AV | `Normal Quantity` | Portion of submitted quantity accepted as normal. |
| AW | `Duplicate Quantity` | Portion of submitted quantity treated as duplicate. |
| AX | `EPMS Evidence` | EPMS row, DU, Tx SOW, and region used as evidence. |
| AY | `PR Model Evidence` | PR Model worksheet row references or fixed rule used to generate expected items. |
| AZ | `Explanation` | Human-readable explanation for the audit result. |

## Workflow

### 1. Preserve Final PO Source Row

Column `AO` is generated during workbook reading. It stores the original row number from `Final PO.xlsx`.

```text
Final PO row number -> AO Source Row
```

### 2. Infer Scope

Column `AP` is produced from Final PO item code, business domain, and item description.

Current scope rules:

| Condition | Scope |
|---|---|
| Item code `350001000403` | `PLANNING` |
| Item code `350000592793` | `OPERATION` |
| Survey/TSS domain text | `TSS` |
| Microwave/TI/antenna domain text or known antenna item code | `TI` |
| No supported mapping | `UNKNOWN` |

### 3. Match EPMS by DU

The auditor matches Final PO to EPMS using DU:

```text
Final PO 逻辑站点编码
=
EPMS du code
```

This match produces `AX EPMS Evidence`.

Example:

```text
EPMS row 5; DU=DU00005252305; Tx SOW=MW New Link / Reroute; Region=Northern
```

### 4. Determine Expected Subcontractor

Column `AU` is selected by scope:

| Scope | Expected subcontractor source |
|---|---|
| `TSS` | EPMS `SubCon - TSS Team` |
| `TI` | EPMS `SubCon - TI Team` |
| `PLANNING` | EPMS `Subcon - Planning` |
| `OPERATION` | Fixed rule: `Allstar` |

Subcontractor mismatch has highest audit priority.

### 5. Generate Expected Item

Columns `AS`, `AT`, and `AY` come from expected entitlement generation.

`AS Expected Item` follows the same item-code selection pattern used by `create-pr-cd` when generating ECC PBOM codes. The auditor must not dump every broad PR Model candidate into AS.

Current baseline rules:

| Scope | Expected item behavior |
|---|---|
| `PLANNING` | Fixed item `350001000403`, quantity `1`. |
| `OPERATION` | Fixed item `350000592793`, quantity `1`, if operation trigger exists. |
| `TSS` | Mandatory PR Model rows matched by Tx SOW, same as create-pr-cd TSS matching. |
| `TI` | Mandatory PR Model rows matched by Tx SOW, then narrowed by create-pr-cd choose-group logic. |

`AY PR Model Evidence` records either a fixed rule or worksheet row references.

Example:

```text
TX Line Item (After 21-Apr 26)!253
```

For example, when EPMS Tx SOW is `MW New Link / Reroute` and antenna size resolves to `1.2m`, AS should show the selected create-pr-cd PBOM item:

```text
350001095410
```

It should not show all `MW New Link / Reroute` mandatory transport, packing, antenna, and dismantle candidates.

Current TI selection behavior:

1. Collect mandatory TI PR Model rows whose SOW matches EPMS `Tx SOW`.
2. Exclude MW Hardware Cutover optional/manual rows.
3. Group duplicate/choose candidates like create-pr-cd:
   - non-choose rows by item code
   - choose rows by SOW, rule text, and normalized choose category
4. Resolve antenna choose groups by selected antenna size.
5. Resolve region/material choose groups only when a single deterministic match exists.
6. Fail closed when the create-pr-cd route/geography resolver is required but not yet wired into the auditor.

### 6. Classify the Row

Columns `AQ`, `AR`, and `AZ` are produced by the audit engine.

Decision priority:

```text
Subcontractor mismatch
-> Invalid domain or missing trigger
-> Wrong item mapping
-> Quantity and duplicate check
-> Normal
```

Common outputs:

| Audit Result | Reason Code Example | Meaning |
|---|---|---|
| `Normal` | `NORMAL_FULL` | Submitted item and quantity match expected entitlement. |
| `Abnormal - Invalid PO` | `INVALID_SUBCON_CHANGED` | Submitted subcontractor differs from expected subcontractor. |
| `Abnormal - Invalid PO` | `INVALID_WRONG_DOMAIN` | Submitted item/domain cannot be mapped to a supported audit scope. |
| `Abnormal - Wrong PO` | `WRONG_LINE_ITEM_MAPPING` | Submitted item is auditable but not one of the expected items. |
| `Abnormal - Duplicate PO` | `DUPLICATE_FULL_QUANTITY` | Valid claim exceeds already-consumed expected quantity. |
| `Abnormal - Duplicate PO` | `DUPLICATE_PARTIAL_QUANTITY` | Part of the submitted quantity is normal and part is duplicate. |

### 7. Resolve Quantity and Duplicates

Columns `AV` and `AW` are populated after audit validation.

Only otherwise-valid rows enter duplicate resolution. Invalid PO and Wrong PO rows do not consume quantity.

Valid claims consume expected quantity in this order:

1. Dispatch Date
2. Request Number
3. Dispatch Order Number
4. PO Line Number

Consumption key:

```text
DU + Scope + Submitted Item Code
```

Quantity outcomes:

| Condition | AQ Audit Result | AV Normal Quantity | AW Duplicate Quantity |
|---|---|---:|---:|
| Valid quantity is within remaining entitlement | `Normal` | submitted quantity | 0 |
| Valid quantity exceeds entitlement | `Abnormal - Duplicate PO` | available portion | excess portion |
| Wrong or Invalid row | Wrong/Invalid classification | 0 | 0 |

## Code References

The implementation lives in:

```text
.openclaw/skills/tx-pr-auditor/scripts/audit_final_po.py
```

Important functions:

| Function | Role |
|---|---|
| `report_writer` | Creates AO-AZ columns in `PR_Audit_Result.xlsx`. |
| `infer_scope` | Produces `AP Scope`. |
| `epms_matcher` | Matches Final PO to EPMS and prepares expected subcontractor. |
| `expected_item_generator` | Produces expected items and PR Model evidence. |
| `audit_engine` | Produces classification, reason code, and explanation. |
| `duplicate_resolver` | Produces normal and duplicate quantities. |
