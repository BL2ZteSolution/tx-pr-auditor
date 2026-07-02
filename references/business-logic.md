# TX PR Auditor Business Logic

This reference summarizes `TX_PR_Auditor_Business_Logic_Specification_v1.0`. Treat that source document as the governing business baseline.

## Scope

Audit only submitted PR/PO lines present in the current `Final PO.xlsx` snapshot. Do not report missing PR/PO, historical deductions, cancelled records that disappeared from the export, claimability, photo quality, or work acceptance quality.

Source of truth:

- Final PO: submitted PR/PO population and submitted line details.
- EPMS: expected operational facts and expected subcontractors.
- PR Model: expected item, mapping, rule, worksheet, and effective-date behavior.

Do not use PR number as the primary audit key. Use DU-level business activity.

## Classification Priority

Apply rules in this order:

1. `Abnormal - Invalid PO`
2. `Abnormal - Wrong PO`
3. `Abnormal - Duplicate PO`
4. `Normal`

Subcontractor mismatch has highest priority. If submitted subcontractor does not match expected subcontractor, classify `Abnormal - Invalid PO` with `INVALID_SUBCON_CHANGED` and stop.

## Master Flow

For each Final PO row:

1. Identify DU, submitted scope/domain, submitted item code, submitted quantity, and ordering fields.
2. Match DU to EPMS.
3. Determine expected subcontractor from EPMS. For Planning, use `Subcon - Planning`.
4. If submitted subcontractor differs from expected subcontractor, return Invalid PO.
5. Determine whether submitted item belongs to the applicable business domain for the DU.
6. If submitted item belongs to a different or unsupported domain, return Invalid PO.
7. Generate expected item set from EPMS facts and applicable PR Model worksheet.
8. If submitted item is in the same domain but does not match expected mapping, return Wrong PO.
9. If submitted item is correct, compare quantity with remaining expected quantity in deterministic snapshot order.
10. Return Normal for available quantity and Duplicate PO for excess quantity.

## Invalid PO

Use `Abnormal - Invalid PO` when:

- Submitted subcontractor differs from expected subcontractor.
- Submitted item belongs to a different implementation domain.
- No supporting EPMS business fact exists for the claimed activity.
- Operation item is submitted while Integration End Date is blank.
- Submitted item is not applicable to the DU scope.

Invalid PO consumes no expected quantity.

## Wrong PO

Use `Abnormal - Wrong PO` when the business activity exists, subcontractor is correct, and submitted item belongs to the same business domain, but expected mapping is wrong.

Examples:

- Wrong antenna-size bracket.
- Wrong regional item.
- Wrong configuration item.
- Wrong line-item mapping.
- Wrong model version.

Wrong PO consumes no expected quantity.

## Normal and Duplicate

A submitted quantity can be Normal only after all validity checks pass and expected quantity remains.

Duplicate detection is quantity-based, not row-count-based:

```text
Remaining Expected Quantity = Expected Quantity - Previously Consumed Normal Quantity
```

If a valid submitted row is partly normal and partly duplicate, do not split the source row. Classify the row as `Abnormal - Duplicate PO`, preserve original submitted quantity, and populate `Normal Quantity` plus `Duplicate Quantity`.

## Quantity Consumption Order

For otherwise-valid claims, consume quantity in this order:

1. Dispatch Date ascending
2. Request Number ascending
3. Dispatch Order Number ascending
4. PO Line Number ascending

Wrong PO and Invalid PO rows never consume quantity and cannot make later valid rows duplicate.

## Dispatch Status

Audit these current snapshot statuses normally:

- `变更中`
- `已关闭`
- `待回执`
- `执行中`

Do not treat `变更中` as a deduction. If the PO disappears from a future export after cancellation/deduction, it will be outside that future snapshot.

## Scope Rules

Planning:

- Expected item code: `350001000403`
- Expected quantity: `1` per DU
- Expected subcontractor: EPMS `Subcon - Planning`
- No antenna logic
- No regional variation

Operation:

- Expected item code: `350000592793`
- Expected quantity: `1` per DU
- Business trigger: Integration End Date exists in EPMS
- No PO date comparison against Integration End Date
- No antenna logic
- No regional variation

TSS:

- Survey scope
- DU matching unit
- Use TSS item rules
- No TI antenna logic

TI:

- Installation scope
- DU matching unit
- Use implementation mapping, region, antenna size, configuration where applicable, expected subcontractor, and effective rule date.
- Do not rely only on TX SOW text. Multiple TX SOW names can share one mapping group.

TSS, TI, Planning, and Operation are separate audit scopes. Never treat TSS and TI as duplicates of each other.

## Antenna Material Mapping

| Antenna size | Expected material code |
|---|---|
| `<= 0.6m` | `350001095405` |
| `> 0.6m and <= 1.2m` | `350001095406` |
| `> 1.2m and <= 1.8m` | `350001095407` |
| `> 1.8m` | `350001095408` |

For legacy worksheet rules, different NE/FE antenna sizes can produce split expected quantities such as `0.5 + 0.5`.

For newer MAX antenna rules, use the larger of NE and FE antenna size with quantity `1`.

Select the applicable rule using effective date information inside the relevant PR Model worksheet, not the workbook filename alone.

## Region Logic

Normalize states into audit regions:

| State/location | Audit region |
|---|---|
| Johor, Melaka, Malacca, Negeri Sembilan | Southern |
| Selangor, Kuala Lumpur, KL | Central |
| Perak, Kedah, Perlis, Penang, Pulau Pinang | Northern |
| Kelantan, Pahang, Terengganu | Eastern |
| Sabah, Sarawak | East Malaysia |

A regional item mismatch in the correct implementation domain is `Abnormal - Wrong PO` with reason `WRONG_REGION_ITEM`.

## Recommended Reason Codes

Normal:

- `NORMAL_FULL`
- `NORMAL_PARTIAL`

Invalid PO:

- `INVALID_SUBCON_CHANGED`
- `INVALID_WRONG_DOMAIN`
- `INVALID_NO_OPERATION_TRIGGER`
- `INVALID_NO_EXPECTED_BUSINESS_FACT`

Wrong PO:

- `WRONG_REGION_ITEM`
- `WRONG_ANTENNA_SIZE`
- `WRONG_CONFIGURATION`
- `WRONG_LINE_ITEM_MAPPING`
- `WRONG_MODEL_VERSION`

Duplicate PO:

- `DUPLICATE_FULL_QUANTITY`
- `DUPLICATE_PARTIAL_QUANTITY`
- `DUPLICATE_ENTITLEMENT_EXHAUSTED`

## Required Output Evidence

Each audit result should preserve:

- Source row, DU, scope, request number, dispatch order number, PO line number, dispatch date.
- Submitted item code, description, quantity, settlement quantity, subcontractor, dispatch status.
- Expected item code/set, expected quantity, expected subcontractor, expected region, applicable model worksheet.
- Classification, reason code, Normal Quantity, Duplicate Quantity, and explanation.
- Relevant EPMS values and PR Model mapping reference.

## Known Implementation Inputs

These are not policy blockers, but the implementation must either resolve them or fail closed with clear evidence:

- Complete TX SOW to Mapping Group to Expected Line-Item Set matrix.
- Exact worksheet names and effective-date ranges from PR Model workbooks.
- Actual operation/integration EPMS source-column mapping.
- Acceptance test cases for each classification and boundary condition.
