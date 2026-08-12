import hashlib
import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from tests.test_audit_final_po import audit


ROOT = Path(__file__).resolve().parents[1]
SPEC = __import__("importlib.util").util.spec_from_file_location("tx_auditor_integration_contract", ROOT / "src" / "main.py")
contract = __import__("importlib.util").util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(contract)


def first_headers(field_map):
    headers = {}
    for source, canonical in field_map.items():
        headers.setdefault(canonical, source)
    return headers


class SkillContractIntegrationTests(unittest.TestCase):
    def test_real_workbooks_run_through_standalone_contract(self):
        with tempfile.TemporaryDirectory() as temp:
            workspace = Path(temp)
            input_dir = workspace / "input"
            input_dir.mkdir()

            final_headers = first_headers(audit.FINAL_PO_FIELD_MAP)
            final_values = {
                "dispatch_date": datetime(2026, 8, 1),
                "dispatch_order_number": "DO-001",
                "po_line_number": 1,
                "request_number": "REQ-001",
                "project_name": "Malaysia_CelcomDigi_Project",
                "business_domain": "TSS survey",
                "submitted_subcontractor": "GCI",
                "logical_site_name": "Wireless RAN/TX Mini Project/SITE-001",
                "du": "DU-001",
                "site_code": "SITE-001",
                "submitted_item_code": "PBOM-001",
                "submitted_item_description": "Synthetic TSS service",
                "submitted_unit": "Each",
                "submitted_quantity": 1,
                "settlement_quantity": 1,
            }
            final_po = input_dir / "Final_PO.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = audit.FINAL_PO_SHEET_NAME
            sheet.append(list(final_headers.values()))
            sheet.append([final_values.get(field, "") for field in final_headers])
            workbook.save(final_po)

            ecc_headers = first_headers(audit.ECC_FIELD_MAP)
            ecc_values = {
                "site_code": "SITE-001",
                "du": "DU-001",
                "region": "Northern",
                "expected_subcontractor": "GCI",
                "expected_item_code": "PBOM-001",
                "expected_item_description": "Synthetic TSS service",
                "expected_unit": "Each",
                "expected_quantity": 1,
                "du_model_name": "TX Mini Project",
                "du_model_id": "4188808420049567786",
                "du_profile_id": "tx_mini_pr_v1",
                "du_view_id": "2477626672974883536",
            }
            ecc = input_dir / "Northern-GCI TX Mini Project TSS PR 20260801.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = audit.ECC_SHEET_NAME
            sheet.append(list(ecc_headers.values()))
            sheet.append([ecc_values.get(field, "") for field in ecc_headers])
            workbook.save(ecc)

            epms = input_dir / "EPMS.xlsx"
            epms.write_bytes(b"synthetic epms input")
            files = []
            for role, file_path in (("final_po", final_po), ("epms", epms)):
                files.append({
                    "name": role,
                    "path": file_path.relative_to(workspace).as_posix(),
                    "originalFileName": file_path.name,
                    "size": file_path.stat().st_size,
                    "sha256": hashlib.sha256(file_path.read_bytes()).hexdigest(),
                })
            envelope = {
                "schemaVersion": "1.0",
                "jobId": "TX-CONTRACT-INTEGRATION",
                "skill": {"skillId": "tx-pr-auditor", "version": "1.1.0"},
                "parameters": {"annotateEcc": True},
                "files": files,
                "paths": {"workspace": ".", "output": "output", "result": "result.json", "cancellation": "control/cancel.requested"},
            }
            input_manifest = workspace / "input.json"
            input_manifest.write_text(json.dumps(envelope), encoding="utf-8")
            entitlement = [
                {"scope": "TSS", "status": "succeeded", "workbooks": [ecc], "warnings": [], "metrics": {}},
                {"scope": "TI", "status": "succeeded", "workbooks": [], "warnings": [], "metrics": {}},
            ]
            with patch.object(contract, "run_entitlement_generation", return_value=entitlement):
                self.assertEqual(contract.run(input_manifest), 0)
            payload = json.loads((workspace / "result.json").read_text(encoding="utf-8"))
            self.assertEqual(payload["status"], "succeeded")
            self.assertEqual(payload["summary"]["metrics"]["classifications"], {"Normal": 1})
            self.assertGreaterEqual(len(payload["outputs"]), 4)
            for output in payload["outputs"]:
                self.assertTrue((workspace / output["path"]).is_file())


if __name__ == "__main__":
    unittest.main()
