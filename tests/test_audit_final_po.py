import importlib.util
import sys
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
        "du": "DU-SYN-001",
        "business_domain": "Planning",
        "submitted_item_code": audit.PLANNING_ITEM_CODE,
        "submitted_item_description": "Synthetic planning service",
        "submitted_quantity": 1.0,
        "settlement_quantity": 1.0,
        "submitted_subcontractor": "Synthetic Vendor Alpha",
    }
    canonical.update(overrides)
    canonical["dispatch_sort_key"] = audit.dispatch_sort_key(canonical)
    return audit.FinalPORecord(source_row=source_row, raw=dict(canonical), canonical=canonical)


def epms_record(source_row=5, du="DU-SYN-001", **overrides):
    canonical = {
        "du": du,
        "site_code": "SITE-SYN-001",
        "site_name": "Synthetic Hill",
        "epms_region": "Synthetic Region",
        "tx_sow": "Synthetic Planning",
        "tx_cutover_date": None,
        "expected_planning_subcontractor": "Alpha Field Services",
        "expected_tss_subcontractor": "Synthetic Vendor Alpha",
        "expected_ti_subcontractor": "Synthetic Vendor Beta",
    }
    canonical.update(overrides)
    return audit.EPMSRecord(source_row=source_row, raw=dict(canonical), canonical=canonical)


def run_pipeline_for_records(final_records, epms_records, pr_model_rows=None):
    dataset = audit.CanonicalDataset(
        final_po_records=final_records,
        epms_records=epms_records,
        pr_model_rows=pr_model_rows or [],
        metadata={},
    )
    matched = audit.epms_matcher(dataset)
    expected = audit.expected_item_generator(matched)
    audited = audit.audit_engine(expected)
    return audit.duplicate_resolver(audited).results


class TxPrAuditorTests(unittest.TestCase):
    def test_duplicate_resolution_handles_mixed_dispatch_dates(self):
        base = audit.AuditResult(
            final_po=final_record(
                source_row=2,
                dispatch_date=datetime(2026, 1, 1),
                request_number="REQ-SYN-001",
            ),
            scope="PLANNING",
            classification="PENDING_QUANTITY",
            reason_code="PENDING_QUANTITY",
            expected_items=[
                audit.ExpectedItem(audit.PLANNING_ITEM_CODE, 2.0, "PLANNING", "Synthetic rule")
            ],
            expected_subcontractor="SYNTHETIC_VENDOR_ALPHA",
            expected_quantity=2.0,
            normal_quantity=0.0,
            duplicate_quantity=0.0,
            explanation="Pending",
            epms_evidence="Synthetic EPMS row",
            pr_model_evidence="Synthetic rule",
            consumes_quantity=True,
        )
        blank_date = replace(
            base,
            final_po=final_record(
                source_row=3,
                dispatch_date="",
                request_number="REQ-SYN-002",
                dispatch_order_number="DO-SYN-002",
            ),
        )
        malformed_date = replace(
            base,
            final_po=final_record(
                source_row=4,
                dispatch_date="not-a-date",
                request_number="REQ-SYN-003",
                dispatch_order_number="DO-SYN-003",
            ),
        )

        results = audit.duplicate_resolver(audit.AuditDataset([malformed_date, blank_date, base], {})).results

        self.assertEqual([result.reason_code for result in results], [
            "DUPLICATE_FULL_QUANTITY",
            "NORMAL_FULL",
            "NORMAL_FULL",
        ])

    def test_unknown_subcontractor_fails_closed(self):
        results = run_pipeline_for_records(
            [final_record(submitted_subcontractor="Unlisted Synthetic Supplier")],
            [epms_record()],
        )

        self.assertEqual(results[0].classification, "Abnormal - Invalid PO")
        self.assertEqual(results[0].reason_code, "INVALID_SUBCON_UNRECOGNIZED")

    def test_missing_epms_match_fails_closed(self):
        results = run_pipeline_for_records([final_record(du="DU-SYN-MISSING")], [epms_record()])

        self.assertEqual(results[0].classification, "Abnormal - Invalid PO")
        self.assertEqual(results[0].reason_code, "INVALID_NO_EXPECTED_BUSINESS_FACT")

    def test_valid_synthetic_normal_case(self):
        results = run_pipeline_for_records([final_record()], [epms_record()])

        self.assertEqual(results[0].classification, "Normal")
        self.assertEqual(results[0].reason_code, "NORMAL_FULL")
        self.assertEqual(results[0].normal_quantity, 1.0)

    def test_wrong_and_duplicate_classification_behavior(self):
        wrong = final_record(
            source_row=2,
            business_domain="Survey",
            submitted_item_code="TSS-SYN-WRONG",
            submitted_item_description="Synthetic TSS service",
        )
        pr_model_rows = [
            {
                "source_row": 10,
                "section": "TSS",
                "sow": "Synthetic Planning",
                "code": "TSS-SYN-EXPECTED",
                "description": "Expected synthetic TSS service",
                "unit": "Each",
                "quantity": 1.0,
                "rules": "Mandatory",
                "is_mandatory": True,
                "worksheet": "Synthetic Model",
            }
        ]
        duplicate_a = final_record(source_row=3, request_number="REQ-SYN-010")
        duplicate_b = final_record(source_row=4, request_number="REQ-SYN-011")

        wrong_result = run_pipeline_for_records([wrong], [epms_record()], pr_model_rows)[0]
        duplicate_results = run_pipeline_for_records([duplicate_a, duplicate_b], [epms_record()])

        self.assertEqual(wrong_result.classification, "Abnormal - Wrong PO")
        self.assertEqual(wrong_result.reason_code, "WRONG_LINE_ITEM_MAPPING")
        self.assertEqual(duplicate_results[0].classification, "Normal")
        self.assertEqual(duplicate_results[1].classification, "Abnormal - Duplicate PO")
        self.assertEqual(duplicate_results[1].reason_code, "DUPLICATE_FULL_QUANTITY")


if __name__ == "__main__":
    unittest.main()
