# Current TX PR Auditor Inputs

Use this reference for the current local workbook contract. The auditor validates `Final PO.xlsx` after `create-pr-cd` has generated ECC output.

## Paths

The runtime must accept explicit paths from the caller. Do not discover files by scanning directories.

Typical local paths:

```text
skills/tx-pr-auditor/input/Final PO.xlsx
skills/create-pr-cd/output/
skills/tx-pr-auditor/output/PR_Audit_Result.xlsx
```

## Workbook Handling

Final PO:

- Auto-detect worksheet `条目明细` with header row `1`.
- Auto-detect worksheet `Sheet1` with header row `2`.
- Allow callers to override the worksheet and header row explicitly.

Generated ECC:

- Worksheet name: `details`
- Header row: row `1`
- Accept `.xlsx` and `.xlsm` files.
- Accept a directory containing generated ECC workbooks.
- Resolve DU identity from `DU Model ID`, `DU Profile ID`, or `DU Model Name`
  columns when present; otherwise resolve it from the ECC filename.
- Use `config/du_registry.json` as the local nine-DU identity contract.

## Final PO Field Map

Map these Final PO headers into canonical fields. The implementation accepts the real Chinese headers shown here and the older mojibake aliases retained for backward compatibility.

| Final PO header | Canonical field |
|---|---|
| `派工日期` | dispatch_date |
| `派工单号` | dispatch_order_number |
| `PO行号` | po_line_number |
| `需求单号` | request_number |
| `项目名称` | project_name |
| `项目编码` | project_code |
| `业务大类` or `能力大类` | business_domain |
| `施工区域` | region |
| `采购区域` | purchasing_area |
| `分包商` | submitted_subcontractor |
| `逻辑站点名称` | logical_site_name |
| `逻辑站点编码` | du |
| `物理站点名称` | physical_site_name |
| `物理站点编码` | site_code |
| `外包代码` | submitted_item_code |
| `代码名称` | submitted_item_description |
| `量纲` | submitted_unit |
| `派工数量` | submitted_quantity |
| `结算数量` | settlement_quantity |
| `支付数量` | paid_quantity |
| `产品型号_备注` | product_model_remark |
| `派工单状态` | dispatch_status |
| `外包商编码` | subcontractor_code |

## Generated ECC Field Map

Map these generated ECC headers into canonical fields:

| ECC header | Canonical field |
|---|---|
| `SN.` | sn |
| `Purchasing Area*` | purchasing_area |
| `Region*` | region |
| `Site ID*` | site_code |
| `Site Name*` | site_name |
| `Delivery Unit Code*` | du |
| `Logical Site Name` | logical_site_name |
| `Contract Number *` | contract_number |
| `Subcontractor*` | expected_subcontractor |
| `PBOM Code*` | expected_item_code |
| `SOW*` | expected_item_description |
| `Unit*` | expected_unit |
| `Quantity*` | expected_quantity |
| `Remarks` | remarks |
| `DU Model` or `DU Model Name` | du_model_name |
| `DU Model ID` | du_model_id |
| `DU Profile ID` | du_profile_id |
| `DU View ID` or `View ID` | du_view_id |
| `DU Project Key` | du_project_key |

## Supported DU Identities

The registry contains nine unique Project + DU Model identities. The two
`CD consolidation 2023` profiles are different views of the same identity and
must not be counted as separate DUs.

The local registry preserves each profile's lifecycle status and accepted View
IDs. When the sibling `create-pr-cd` source contract is available, registry
loading compares the complete identity/profile/view mapping and fails if the
local snapshot has drifted.

If Final PO identifies a DU model in Project Name, product-model remarks, or
logical-site name, matching is restricted to that DU. Project Code is resolved
to the registered create-pr-cd project and must agree with the DU model's
project. If the same submitted site and item match multiple identified DU
models without a Final PO DU model, the row fails closed with
`INVALID_AMBIGUOUS_DU_MODEL`.

## Post-create-pr-cd Validation Role

The auditor validates submitted PO rows in `Final PO.xlsx` against generated ECC rows. It does not generate ECC files, does not call `create-pr-cd`, and does not read EPMS or PR Model workbooks.
