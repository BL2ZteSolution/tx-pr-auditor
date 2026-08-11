import hashlib
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("tx_auditor_contract", ROOT / "src" / "main.py")
contract = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(contract)


class SkillContractTests(unittest.TestCase):
    def make_request(self, root: Path, *, cancelled: bool = False) -> Path:
        input_dir = root / "input"
        input_dir.mkdir()
        files = []
        for name, role in (("final.xlsx", "final_po"), ("ecc.xlsx", "expected_ecc")):
            path = input_dir / name
            path.write_bytes(name.encode())
            files.append({
                "name": role,
                "path": f"input/{name}",
                "size": path.stat().st_size,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            })
        if cancelled:
            (root / "control").mkdir()
            (root / "control" / "cancel.requested").write_text("cancel", encoding="utf-8")
        envelope = {
            "schemaVersion": "1.0",
            "jobId": "JOB-AUDIT-001",
            "skill": {"skillId": "tx-pr-auditor", "version": "1.0.0"},
            "parameters": {"annotateEcc": False},
            "files": files,
            "paths": {"workspace": ".", "output": "output", "result": "result.json", "cancellation": "control/cancel.requested"},
        }
        manifest = root / "input.json"
        manifest.write_text(json.dumps(envelope), encoding="utf-8")
        return manifest

    def test_success_writes_audit_outputs_and_safe_metrics(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_request(root)

            def fake_run(parsed):
                Path(parsed.output).write_bytes(b"audit")
                Path(parsed.summary_json).write_text("{}", encoding="utf-8")
                return {"total_rows": 2, "classifications": {"Normal": 2}, "reason_codes": {}, "du_models": {"DU": 2}}

            with patch.object(contract.audit_final_po, "run_pipeline", side_effect=fake_run):
                self.assertEqual(contract.run(manifest), 0)
            result = json.loads((root / "result.json").read_text(encoding="utf-8"))
            self.assertEqual(result["status"], "succeeded")
            self.assertEqual(result["summary"]["metrics"]["totalRows"], 2)
            self.assertEqual(len(result["outputs"]), 2)

    def test_findings_produce_warning_status(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_request(root)

            def fake_run(parsed):
                Path(parsed.output).write_bytes(b"audit")
                Path(parsed.summary_json).write_text("{}", encoding="utf-8")
                return {"total_rows": 1, "classifications": {"Abnormal - Wrong": 1}}

            with patch.object(contract.audit_final_po, "run_pipeline", side_effect=fake_run):
                self.assertEqual(contract.run(manifest), 0)
            result = json.loads((root / "result.json").read_text(encoding="utf-8"))
            self.assertEqual(result["status"], "succeeded_with_warning")
            self.assertEqual(result["warnings"][0]["code"], "AUDIT_FINDINGS_PRESENT")

    def test_cancellation_stops_before_workbook_processing(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_request(root, cancelled=True)
            with patch.object(contract.audit_final_po, "run_pipeline") as pipeline:
                self.assertEqual(contract.run(manifest), 130)
                pipeline.assert_not_called()
            result = json.loads((root / "result.json").read_text(encoding="utf-8"))
            self.assertEqual(result["status"], "cancelled")


if __name__ == "__main__":
    unittest.main()
