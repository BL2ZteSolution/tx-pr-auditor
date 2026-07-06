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

- Worksheet name: `æ¡ç›®æ˜Žç»†`
- Header row: row `1`

Generated ECC:

- Worksheet name: `details`
- Header row: row `1`
- Accept `.xlsx` and `.xlsm` files.
- Accept a directory containing generated ECC workbooks.

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

## Post-create-pr-cd Validation Role

The auditor validates submitted PO rows in `Final PO.xlsx` against generated ECC rows. It does not generate ECC files, does not call `create-pr-cd`, and does not read EPMS or PR Model workbooks.
