#!/usr/bin/env python3
"""AI Worker Platform contract entrypoint for the standalone TX PR Auditor."""
from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import signal
import sys
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL_ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import audit_final_po  # noqa: E402


CONTRACT_VERSION = "1.0"
SKILL_ID = "tx-pr-auditor"
SKILL_VERSION = "1.0.0"


class ContractError(Exception):
    def __init__(self, code: str, message: str, category: str = "domain_input", details: dict | None = None):
        super().__init__(message)
        self.code = code
        self.category = category
        self.details = details or {}


class CancelledError(ContractError):
    def __init__(self):
        super().__init__("SKILL_CANCELLED", "Cancellation was requested.", "cancelled")


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def emit(event_type: str, phase: str, message: str, percent: int | None = None) -> None:
    event: dict[str, Any] = {"type": event_type, "timestamp": utc_now(), "phase": phase, "message": message}
    if percent is not None:
        event["percent"] = percent
    print(json.dumps(event, ensure_ascii=False), flush=True)


def start_progress_heartbeat(phase: str, message: str, seconds: int = 30):
    stopped = threading.Event()
    def heartbeat() -> None:
        while not stopped.wait(seconds):
            emit("progress", phase, message)
    thread = threading.Thread(target=heartbeat, daemon=True)
    thread.start()
    return stopped


def parse_cli() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run tx-pr-auditor using the AI Worker Platform skill contract.")
    parser.add_argument("--input-manifest", required=True, type=Path)
    return parser.parse_args()


def resolve_inside(workspace: Path, value: Any, label: str) -> Path:
    raw = str(value or "").strip()
    if not raw or Path(raw).is_absolute():
        raise ContractError("CONTRACT_PATH_INVALID", f"{label} must be a workspace-relative path.")
    resolved = (workspace / raw).resolve()
    try:
        resolved.relative_to(workspace)
    except ValueError as exc:
        raise ContractError("CONTRACT_PATH_INVALID", f"{label} escapes the workspace.") from exc
    return resolved


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_envelope(path: Path) -> tuple[dict[str, Any], Path, Path, Path, Path]:
    manifest_path = path.resolve()
    if not manifest_path.is_file():
        raise ContractError("INPUT_MANIFEST_NOT_FOUND", "The input manifest was not found.")
    try:
        envelope = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ContractError("INPUT_MANIFEST_INVALID", "The input manifest is not valid JSON.") from exc
    if envelope.get("schemaVersion") != CONTRACT_VERSION:
        raise ContractError("CONTRACT_VERSION_UNSUPPORTED", "Unsupported input contract version.")
    skill = envelope.get("skill") or {}
    if skill.get("skillId") != SKILL_ID or skill.get("version") != SKILL_VERSION:
        raise ContractError("SKILL_IDENTITY_MISMATCH", "Input manifest skill identity does not match this package.")
    if not str(envelope.get("jobId") or "").strip():
        raise ContractError("JOB_ID_REQUIRED", "jobId is required.")
    paths = envelope.get("paths") or {}
    workspace = resolve_inside(manifest_path.parent, paths.get("workspace", "."), "paths.workspace")
    output = resolve_inside(workspace, paths.get("output", "output"), "paths.output")
    result = resolve_inside(workspace, paths.get("result", "result.json"), "paths.result")
    cancellation = resolve_inside(workspace, paths.get("cancellation", "control/cancel.requested"), "paths.cancellation")
    output.mkdir(parents=True, exist_ok=True)
    result.parent.mkdir(parents=True, exist_ok=True)
    return envelope, workspace, output, result, cancellation


def declared_files(envelope: dict[str, Any], workspace: Path, name: str, multiple: bool) -> list[Path]:
    matches = [item for item in envelope.get("files", []) if item.get("name") == name]
    if not matches or (not multiple and len(matches) != 1):
        quantifier = "one or more" if multiple else "exactly one"
        raise ContractError("INPUT_FILE_INVALID", f"{quantifier} {name} file is required.")
    paths = []
    for index, item in enumerate(matches):
        path = resolve_inside(workspace, item.get("path"), f"files.{name}[{index}].path")
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            raise ContractError("INPUT_FILE_INVALID", f"{name} must contain existing .xlsx files.")
        if item.get("size") is not None and int(item["size"]) != path.stat().st_size:
            raise ContractError("INPUT_FILE_SIZE_MISMATCH", f"{name} size does not match its declaration.")
        if item.get("sha256") and str(item["sha256"]).lower() != sha256(path):
            raise ContractError("INPUT_FILE_CHECKSUM_MISMATCH", f"{name} checksum does not match its declaration.")
        paths.append(path)
    return paths


def check_cancel(path: Path) -> None:
    if path.exists():
        raise CancelledError()


def output_item(path: Path, workspace: Path) -> dict[str, Any]:
    return {
        "name": path.stem,
        "path": path.resolve().relative_to(workspace).as_posix(),
        "mediaType": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
        "displayName": path.name,
        "size": path.stat().st_size,
        "sha256": sha256(path),
    }


def write_result(result_path: Path, payload: dict[str, Any]) -> None:
    temp = result_path.with_suffix(result_path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, result_path)


def safe_metrics(summary: dict[str, Any]) -> dict[str, Any]:
    return {
        "totalRows": int(summary.get("total_rows", 0) or 0),
        "classifications": dict(summary.get("classifications") or {}),
        "reasonCodes": dict(summary.get("reason_codes") or {}),
        "duModels": dict(summary.get("du_models") or {}),
        "annotatedFileCount": int((summary.get("annotated_ecc") or {}).get("file_count", 0) or 0),
    }


def run(input_manifest: Path) -> int:
    envelope: dict[str, Any] = {}
    result_path = input_manifest.resolve().parent / "result.json"
    try:
        envelope, workspace, output, result_path, cancellation = load_envelope(input_manifest)
        signal.signal(signal.SIGTERM, lambda *_: (_ for _ in ()).throw(CancelledError()))
        check_cancel(cancellation)
        emit("progress", "contract_validation", "Validated the audit request.", 5)
        final_po = declared_files(envelope, workspace, "final_po", False)[0]
        expected_ecc = declared_files(envelope, workspace, "expected_ecc", True)
        parameters = envelope.get("parameters") or {}
        allowed = {"filterYear", "filterMonth", "annotateEcc"}
        unknown = sorted(set(parameters) - allowed)
        if unknown:
            raise ContractError("PARAMETERS_INVALID", "Unsupported parameters were supplied.", details={"fields": unknown})
        month = parameters.get("filterMonth")
        year = parameters.get("filterYear")
        if month is not None and (not isinstance(month, int) or not 1 <= month <= 12):
            raise ContractError("PARAMETERS_INVALID", "filterMonth must be an integer from 1 to 12.")
        if year is not None and (not isinstance(year, int) or not 2000 <= year <= 2200):
            raise ContractError("PARAMETERS_INVALID", "filterYear must be an integer from 2000 to 2200.")
        audit_output = output / "PR_Audit_Result.xlsx"
        summary_output = output / "PR_Audit_Summary.json"
        parsed = argparse.Namespace(
            final_po=str(final_po),
            final_po_sheet=None,
            final_po_header_row=None,
            final_po_max_rows=10000,
            filter_year=year,
            filter_month=month,
            expected_ecc=[str(path) for path in expected_ecc],
            du_registry=str(SKILL_ROOT / "config" / "du_registry.json"),
            ecc_sheet=audit_final_po.ECC_SHEET_NAME,
            output=str(audit_output),
            summary_json=str(summary_output),
            annotate_ecc_output=bool(parameters.get("annotateEcc", True)),
            annotated_ecc_output_root=str(output / "annotated-ecc"),
            annotated_ecc_timestamp="current",
        )
        audit_final_po.set_cancellation_probe(lambda: check_cancel(cancellation))
        emit("progress", "workbook_read", "Reading Final PO and ECC workbooks.", 15)
        progress_heartbeat = start_progress_heartbeat("audit_processing", "Workbook audit is still running.")
        try:
            summary = audit_final_po.run_pipeline(parsed)
        finally:
            progress_heartbeat.set()
        check_cancel(cancellation)
        emit("progress", "result_packaging", "Packaging audit outputs.", 95)
        candidates = [audit_output, summary_output]
        annotated = summary.get("annotated_ecc") or {}
        candidates.extend(Path(item["output_file"]) for item in annotated.get("files", []) if item.get("output_file"))
        annotated_summary = Path(annotated["output_dir"]) / "annotated_ecc.summary.json" if annotated.get("output_dir") else None
        if annotated_summary and annotated_summary.is_file():
            candidates.append(annotated_summary)
        candidates = list(dict.fromkeys(path.resolve() for path in candidates if path.is_file()))
        for path in candidates:
            path.relative_to(output.resolve())
        metrics = safe_metrics(summary)
        abnormal = sum(count for name, count in metrics["classifications"].items() if str(name).lower() != "normal")
        status = "succeeded_with_warning" if abnormal else "succeeded"
        payload = {
            "schemaVersion": CONTRACT_VERSION,
            "jobId": envelope["jobId"],
            "skillId": SKILL_ID,
            "skillVersion": SKILL_VERSION,
            "status": status,
            "summary": {"message": "TX PR audit completed.", "metrics": metrics},
            "outputs": [output_item(path, workspace) for path in candidates],
            "warnings": ([{"code": "AUDIT_FINDINGS_PRESENT", "message": "The audit contains findings requiring review.", "details": {"count": abnormal}}] if abnormal else []),
            "error": None,
        }
        write_result(result_path, payload)
        emit("progress", "completed", "TX PR audit completed.", 100)
        return 0
    except CancelledError as exc:
        status, exit_code, error = "cancelled", 130, exc
    except ContractError as exc:
        status, exit_code, error = "failed", 2, exc
    except Exception as exc:
        status, exit_code = "failed", 4
        error = ContractError(getattr(exc, "code", "TX_PR_AUDIT_FAILED"), str(exc), "domain_processing")

    payload = {
        "schemaVersion": CONTRACT_VERSION,
        "jobId": str(envelope.get("jobId") or "unknown"),
        "skillId": SKILL_ID,
        "skillVersion": SKILL_VERSION,
        "status": status,
        "summary": {"message": str(error), "metrics": {}},
        "outputs": [],
        "warnings": [],
        "error": {"code": error.code, "category": error.category, "message": str(error), "retryable": False, "details": error.details},
    }
    result_path.parent.mkdir(parents=True, exist_ok=True)
    write_result(result_path, payload)
    emit("warning", status, str(error))
    return exit_code


def main() -> int:
    args = parse_cli()
    return run(args.input_manifest)


if __name__ == "__main__":
    raise SystemExit(main())
