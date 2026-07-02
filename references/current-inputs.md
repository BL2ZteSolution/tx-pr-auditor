# Current TX PR Auditor Inputs

Use this reference for the current local workbook contract. The auditor validates `Final PO.xlsx` after `create-pr-cd` work has produced or refreshed the submitted PO population.

## Paths

Current fixture paths:

```text
.openclaw/skills/tx-pr-auditor/input/Final PO.xlsx
.openclaw/skills/tx-pr-auditor/input/EPMS.xlsx
.openclaw/skills/tx-pr-auditor/input/pr_model.xlsx
```

The runtime must still accept explicit paths from the caller. Do not discover files by scanning directories.

## Workbook Handling

Final PO:

- Use worksheet name `条目明细`.
- Header row: row `1`.
- Current dimension observed: `A1:AN1000`.

EPMS:

- Use worksheet name `data`.
- Machine/source headers occupy rows `1` to `3`.
- Human-readable implementation headers are on row `4`.
- Data starts on row `5`.
- Current dimension observed: `A1:EL1000`.

PR Model:

- Use the supplied workbook only.
- Current filename is `pr_model.xlsx`.
- Current primary sheet observed: `TX Line Item (After 21-Apr 26)`.
- The implementation must select applicable worksheet/rules by worksheet effective date metadata where available, not by filename alone.

## Final PO Field Map

Map these Chinese Final PO headers into canonical fields:

| Final PO header | Canonical field |
|---|---|
| `派工日期` | dispatch_date |
| `派工单号` | dispatch_order_number |
| `PO行号` | po_line_number |
| `需求单号` | request_number |
| `项目名称` | project_name |
| `项目编码` | project_code |
| `业务大类` | business_domain |
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

## EPMS Field Map

Read EPMS using row `4` as the human-readable header row.

Key canonical fields:

| EPMS header | Canonical field |
|---|---|
| `customer site code` | site_code |
| `customer site name` | site_name |
| `du code` | du |
| `region` | epms_region |
| `Province/State` | province_state |
| `Latitude (North Plus South Minus)` | latitude |
| `Longitude (East Plus West Minus)` | longitude |
| `TX Upgrade Scope` | tx_upgrade_scope |
| `BOQ Configuration` | boq_configuration |
| `Tx SOW` | tx_sow |
| `TX SOW Details` | tx_sow_details |
| `NE SOW Details` | ne_sow_details |
| `FE SOW Details` | fe_sow_details |
| `MW Config Antenna Size NE` | antenna_size_ne |
| `MW Config Antenna Size FE` | antenna_size_fe |
| `SubCon - TSS Team` | expected_tss_subcontractor |
| `Subcon PR - TSS` | existing_tss_pr |
| `SubCon - TI Team` | expected_ti_subcontractor |
| `Subcon PR - TI` | existing_ti_pr |
| `TX Cutover Date` | tx_cutover_date |
| `Subcon - Planning` | expected_planning_subcontractor |
| `Subcon PR - Planning` | existing_planning_pr |

Operation/integration trigger fields must be confirmed during implementation because EPMS contains repeated `actual end time` columns under grouped source headers. Preserve source column address/evidence when mapping these fields.

## Post-create-pr-cd Validation Role

The auditor validates submitted PO rows in `Final PO.xlsx`. It does not generate ECC files and must not call or modify `create-pr-cd/scripts/generate_tss_pr_ecc.py`.

Use `create-pr-cd` knowledge only as reference context for PR model behavior and field semantics. Keep audit implementation independent.
