import importlib.util
import sys
import tempfile
import unittest
from dataclasses import replace
from datetime import datetime
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "audit_final_po.py"
SPEC = importlib.util.spec_from_file_location("audit_final_po", SCRIPT_PATH)
audit = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = audit
SPEC.loader.exec_module(audit)


def final_record(source_row=2, **overrides):
    canonical = {
        "dispatch_date": datetime(2026, 1, 2),
        "request_number": f"REQ-SYN-{source_row:03d}",
        "dispatch_order_number": f"DO-SYN-{source_row:03d}",
        "po_line_number": source_row,
        "site_code": "SITE-SYN-001",
        "du": "DU-SYN-001",
        "business_domain": "Planning",
        "submitted_item_code": "350001000403",
        "submitted_item_description": "Synthetic planning service",
        "submitted_quantity": 1.0,
        "settlement_quantity": 1.0,
        "submitted_subcontractor": "GCI",
    }
    canonical.update(overrides)
    canonical["site_code"] = audit.normalize_code(canonical.get("site_code"))
    canonical["du"] = audit.normalize_code(canonical.get("du"))
    canonical["submitted_item_code"] = audit.normalize_code(canonical.get("submitted_item_code"))
    canonical["submitted_subcontractor_norm"] = audit.normalize_subcontractor(canonical.get("submitted_subcontractor"))
    canonical["dispatch_sort_key"] = audit.dispatch_sort_key(canonical)
    return audit.FinalPORecord(source_row=source_row, raw=dict(canonical), canonical=canonical)


def expected_record(source_row=2, **overrides):
    source_file = overrides.pop("source_file", "Northern-GCI TX Mini Project Planning PR 20260706.xlsx")
    source_sheet = overrides.pop("source_sheet", "details")
    canonical = {
        "site_code": "SITE-SYN-001",
        "du": "DU-SYN-001",
        "region": "Northern",
        "expected_subcontractor": "GCI",
        "expected_item_code": "350001000403",
        "expected_item_description": "Synthetic planning service",
        "expected_quantity": 1.0,
        "scope": "PLANNING",
    }
    canonical.update(overrides)
    canonical["site_code"] = audit.normalize_code(canonical.get("site_code"))
    canonical["du"] = audit.normalize_code(canonical.get("du"))
    canonical["expected_item_code"] = audit.normalize_code(canonical.get("expected_item_code"))
    canonical["expected_quantity"] = audit.to_float(canonical.get("expected_quantity"))
    canonical["expected_subcontractor_norm"] = audit.normalize_subcontractor(canonical.get("expected_subcontractor"))
    canonical["entitlement_key"] = audit.entitlement_key(canonical)
    return audit.ExpectedECCRecord(
        source_file=source_file,
        source_sheet=source_sheet,
        source_row=source_row,
        raw=dict(canonical),
        canonical=canonical,
    )


def run_pipeline_for_records(final_records, expected_records):
    dataset = audit.CanonicalDataset(final_records, expected_records, metadata={})
    matches = audit.expected_matcher(dataset)
    audited = audit.audit_engine(matches, dataset.metadata)
    return audit.duplicate_resolver(audited).results


class TxPrAuditorTests(unittest.TestCase):
    def test_final_po_layout_auto_detects_supported_formats(self):
        try:
            from openpyxl import Workbook
        except ModuleNotFoundError:
            self.skipTest("openpyxl is required for workbook layout test")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            for sheet_name, header_row in (("条目明细", 1), ("Sheet1", 2)):
                workbook_path = tmp_path / f"{sheet_name}.xlsx"
                wb = Workbook()
                ws = wb.active
                ws.title = sheet_name
                if header_row == 2:
                    ws.cell(1, 36, "PM/TL to feedback")
                ws.cell(header_row, 1, "派工日期")
                ws.cell(header_row, 2, "派工单号")
                ws.cell(header_row + 1, 1, datetime(2026, 1, 2))
                ws.cell(header_row + 1, 2, "DO-SYN-001")
                wb.save(workbook_path)

                resolved = audit.resolve_final_po_layout(workbook_path, None, None)
                rows, metadata = audit.read_table(workbook_path, *resolved)

                self.assertEqual(resolved, (sheet_name, header_row))
                self.assertEqual(metadata["sheet"], sheet_name)
                self.assertEqual(metadata["header_row"], header_row)
                self.assertEqual(rows[0]["派工单号"], "DO-SYN-001")

    def test_final_po_layout_preserves_explicit_overrides(self):
        try:
            from openpyxl import Workbook
        except ModuleNotFoundError:
            self.skipTest("openpyxl is required for workbook layout test")

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = Path(tmpdir) / "custom.xlsx"
            wb = Workbook()
            wb.active.title = "Custom"
            wb.save(workbook_path)

            self.assertEqual(
                audit.resolve_final_po_layout(workbook_path, "Custom", 3),
                ("Custom", 3),
            )

    def test_valid_row_is_normal_against_create_pr_cd_output(self):
        results = run_pipeline_for_records([final_record()], [expected_record()])

        self.assertEqual(results[0].classification, "Normal")
        self.assertEqual(results[0].reason_code, "NORMAL_FULL")
        self.assertEqual(results[0].normal_quantity, 1.0)

    def test_wrong_item_when_site_exists_but_pbom_not_expected(self):
        results = run_pipeline_for_records(
            [final_record(submitted_item_code="WRONG-CODE")],
            [expected_record()],
        )

        self.assertEqual(results[0].classification, "Abnormal - Wrong PO")
        self.assertEqual(results[0].reason_code, "WRONG_LINE_ITEM_MAPPING")

    def test_missing_generated_ecc_entitlement_is_invalid(self):
        results = run_pipeline_for_records(
            [final_record(site_code="SITE-NOT-GENERATED", du="DU-NOT-GENERATED")],
            [expected_record()],
        )

        self.assertEqual(results[0].classification, "Abnormal - Invalid PO")
        self.assertEqual(results[0].reason_code, "INVALID_NOT_IN_CREATE_PR_CD_OUTPUT")

    def test_subcontractor_changed_is_invalid(self):
        results = run_pipeline_for_records(
            [final_record(submitted_subcontractor="Other Supplier")],
            [expected_record()],
        )

        self.assertEqual(results[0].classification, "Abnormal - Invalid PO")
        self.assertEqual(results[0].reason_code, "INVALID_SUBCON_CHANGED")

    def test_duplicate_consumption_uses_dispatch_order(self):
        first = final_record(source_row=2, dispatch_date=datetime(2026, 1, 1), request_number="REQ-001")
        second = final_record(source_row=3, dispatch_date=datetime(2026, 1, 2), request_number="REQ-002")

        results = run_pipeline_for_records([second, first], [expected_record(expected_quantity=1.0)])

        self.assertEqual(results[0].classification, "Abnormal - Duplicate PO")
        self.assertEqual(results[0].reason_code, "DUPLICATE_FULL_QUANTITY")
        self.assertEqual(results[1].classification, "Normal")

    def test_partial_duplicate_preserves_normal_and_duplicate_quantities(self):
        submitted = final_record(submitted_quantity=2.0, settlement_quantity=2.0)
        results = run_pipeline_for_records([submitted], [expected_record(expected_quantity=1.5)])

        self.assertEqual(results[0].classification, "Abnormal - Duplicate PO")
        self.assertEqual(results[0].reason_code, "DUPLICATE_PARTIAL_QUANTITY")
        self.assertEqual(results[0].normal_quantity, 1.5)
        self.assertEqual(results[0].duplicate_quantity, 0.5)

    def test_annotation_status_aggregation(self):
        normal = run_pipeline_for_records([final_record()], [expected_record()])[0]
        duplicate = run_pipeline_for_records(
            [
                final_record(source_row=2, request_number="REQ-001"),
                final_record(source_row=3, request_number="REQ-002"),
            ],
            [expected_record(expected_quantity=1.0)],
        )[1]

        self.assertEqual(audit.annotation_values([normal])[0], "NORMAL")
        self.assertEqual(audit.annotation_values([])[0], "NOT_IN_FINAL_PO")
        self.assertEqual(audit.annotation_values([duplicate])[0], "DUPLICATE")
        self.assertEqual(audit.annotation_values([normal, duplicate])[0], "MIXED")

    def test_annotated_ecc_writer_copies_without_touching_source(self):
        try:
            from openpyxl import Workbook, load_workbook
        except ModuleNotFoundError:
            self.skipTest("openpyxl is required for workbook annotation test")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            source = tmp_path / "Northern-GCI TX Mini Project Planning PR 20260706.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "details"
            headers = [
                "SN.",
                "Purchasing Area*",
                "Region*",
                "Site ID*",
                "Site Name*",
                "Delivery Unit Code*",
                "Logical Site Name",
                "Contract Number *",
                "Subcontractor*",
                "PBOM Code*",
                "SOW*",
                "Unit*",
                "Quantity*",
                "Remarks",
                None,
                "Contract Number",
            ]
            for idx, header in enumerate(headers, 1):
                ws.cell(1, idx, header)
            ws.cell(2, 1, 1)
            ws.cell(2, 4, "SITE-SYN-001")
            ws.cell(2, 6, "DU-SYN-001")
            ws.cell(2, 9, "GCI")
            ws.cell(2, 10, "350001000403")
            ws.cell(2, 13, 1)
            wb.save(source)

            result = run_pipeline_for_records(
                [final_record()],
                [expected_record(source_file=str(source), source_row=2)],
            )[0]
            summary = audit.annotated_ecc_writer(
                audit.AuditDataset([result], {}),
                [source],
                tmp_path / "output",
                "TEST_RUN",
                "details",
            )

            source_wb = load_workbook(source, read_only=True, data_only=True)
            source_headers = [cell.value for cell in next(source_wb["details"].iter_rows(min_row=1, max_row=1))]
            self.assertNotIn("Audit Status", source_headers)
            source_wb.close()

            copied = tmp_path / "output" / "TEST_RUN" / source.name
            copied_wb = load_workbook(copied, read_only=True, data_only=True)
            copied_ws = copied_wb["details"]
            copied_headers = [cell.value for cell in next(copied_ws.iter_rows(min_row=1, max_row=1))]
            self.assertIn("Audit Status", copied_headers)
            self.assertEqual(copied_ws.cell(2, 17).value, "NORMAL")
            self.assertEqual(Path(summary["files"][0]["output_file"]).name, source.name)
            copied_wb.close()


if __name__ == "__main__":
    unittest.main()
