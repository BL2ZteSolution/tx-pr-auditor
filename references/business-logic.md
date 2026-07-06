# TX PR Auditor Business Logic

## Scope

Audit only submitted PR/PO lines present in the current `Final PO.xlsx` snapshot.

Source of truth:

- Generated ECC output from `create-pr-cd`: expected entitlement.
- Final PO: submitted PR/PO population and submitted line details.

Do not use EPMS, PR Model, or PR number as audit truth. The auditor validates whether Final PO matches what `create-pr-cd` generated.

## Classification Priority

Apply rules in this order:

1. `Abnormal - Invalid PO`
2. `Abnormal - Wrong PO`
3. `Abnormal - Duplicate PO`
4. `Normal`

Subcontractor mismatch has high priority. If submitted subcontractor does not match generated ECC subcontractor for the site or item, classify `Abnormal - Invalid PO` with `INVALID_SUBCON_CHANGED`.

## Master Flow

For each Final PO row:

1. Identify submitted site code, DU, item code, subcontractor, quantity, and ordering fields.
2. Match site code or DU to generated ECC rows.
3. If no generated ECC entitlement exists, return Invalid PO.
4. Validate subcontractor against generated ECC subcontractor.
5. Check whether submitted item code exists in generated ECC rows for the site.
6. If site exists but item does not, return Wrong PO.
7. If item is correct, compare quantity with remaining generated ECC entitlement in deterministic snapshot order.
8. Return Normal for available quantity and Duplicate PO for excess quantity.

## Invalid PO

Use `Abnormal - Invalid PO` when:

- No generated ECC entitlement exists for the submitted site or DU.
- Submitted subcontractor differs from generated ECC subcontractor.

Invalid PO consumes no expected quantity.

## Wrong PO

Use `Abnormal - Wrong PO` when:

- The submitted site or DU exists in generated ECC output.
- The submitted subcontractor is acceptable.
- The submitted PBOM/item code is not expected for that site or DU.

Wrong PO consumes no expected quantity.

## Normal and Duplicate

A submitted quantity can be Normal only after all validity checks pass and expected quantity remains.

Duplicate detection is quantity-based:

```text
Remaining Expected Quantity = Generated ECC Quantity - Previously Consumed Normal Quantity
```

If a valid submitted row is partly normal and partly duplicate, do not split the source row. Classify the row as `Abnormal - Duplicate PO`, preserve original submitted quantity, and populate `Normal Quantity` plus `Duplicate Quantity`.

## Quantity Consumption Order

For otherwise-valid claims, consume quantity in this order:

1. Dispatch Date ascending
2. Request Number ascending
3. Dispatch Order Number ascending
4. PO Line Number ascending

Wrong PO and Invalid PO rows never consume quantity and cannot make later valid rows duplicate.

## Recommended Reason Codes

Normal:

- `NORMAL_FULL`

Invalid PO:

- `INVALID_NOT_IN_CREATE_PR_CD_OUTPUT`
- `INVALID_SUBCON_CHANGED`

Wrong PO:

- `WRONG_LINE_ITEM_MAPPING`

Duplicate PO:

- `DUPLICATE_FULL_QUANTITY`
- `DUPLICATE_PARTIAL_QUANTITY`

Internal transitional state:

- `PENDING_QUANTITY`

## Required Output Evidence

Each audit result should preserve:

- Source row, site code, DU, scope, request number, dispatch order number, PO line number, dispatch date.
- Submitted item code, description, quantity, settlement quantity, subcontractor, dispatch status.
- Expected item code/set, expected quantity, expected subcontractor.
- Generated ECC file name, worksheet, and row number evidence.
- Classification, reason code, Normal Quantity, Duplicate Quantity, and explanation.
