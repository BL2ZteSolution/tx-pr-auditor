#!/usr/bin/env python3
"""AI Worker Platform contract entrypoint for the standalone TX PR Auditor."""
from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import os
import queue
import shutil
import signal
import subprocess
import sys
import threading
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SKILL_ROOT = Path(__file__).resolve().parents[1]
CREATE_PR_CD_ROOT = SKILL_ROOT / "dependencies" / "create-pr-cd"
SCRIPTS = SKILL_ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import audit_final_po  # noqa: E402


CONTRACT_VERSION = "1.0"
SKILL_ID = "tx-pr-auditor"
SKILL_VERSION = "1.1.0"
CREATE_PR_CD_SKILL_ID = "create-pr-cd"
CREATE_PR_CD_VERSION = "4.0.0"
ENTITLEMENT_SCOPES = ("TSS", "TI")


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


def declared_files(
    envelope: dict[str, Any],
    workspace: Path,
    name: str,
    multiple: bool,
    extensions: tuple[str, ...] = (".xlsx",),
) -> list[Path]:
    matches = [item for item in envelope.get("files", []) if item.get("name") == name]
    if not matches or (not multiple and len(matches) != 1):
        quantifier = "one or more" if multiple else "exactly one"
        raise ContractError("INPUT_FILE_INVALID", f"{quantifier} {name} file is required.")
    paths = []
    for index, item in enumerate(matches):
        path = resolve_inside(workspace, item.get("path"), f"files.{name}[{index}].path")
        if not path.is_file() or path.suffix.lower() not in extensions:
            accepted = ", ".join(extensions)
            raise ContractError("INPUT_FILE_INVALID", f"{name} must contain existing files with one of: {accepted}.")
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


def create_delivery_archive(output: Path, candidates: list[Path]) -> Path:
    archive = output / "TX_PR_Audit_Delivery.zip"
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        for path in candidates:
            bundle.write(path, path.relative_to(output).as_posix())
    return archive


def load_create_pr_cd_dependency() -> tuple[Path, str]:
    manifest_path = CREATE_PR_CD_ROOT / "skill.json"
    if not manifest_path.is_file():
        raise ContractError(
            "CREATE_PR_CD_DEPENDENCY_MISSING",
            "The pinned create-pr-cd dependency is unavailable. Initialize Git submodules recursively.",
            "dependency",
        )
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ContractError(
            "CREATE_PR_CD_DEPENDENCY_INVALID",
            "The pinned create-pr-cd dependency manifest is invalid.",
            "dependency",
        ) from exc
    if manifest.get("skillId") != CREATE_PR_CD_SKILL_ID or manifest.get("version") != CREATE_PR_CD_VERSION:
        raise ContractError(
            "CREATE_PR_CD_DEPENDENCY_IDENTITY_MISMATCH",
            "The pinned create-pr-cd dependency identity does not match the auditor release.",
            "dependency",
            {
                "expectedSkillId": CREATE_PR_CD_SKILL_ID,
                "expectedVersion": CREATE_PR_CD_VERSION,
                "actualSkillId": manifest.get("skillId"),
                "actualVersion": manifest.get("version"),
            },
        )
    entrypoint = (CREATE_PR_CD_ROOT / str((manifest.get("runtime") or {}).get("entrypoint") or "")).resolve()
    try:
        entrypoint.relative_to(CREATE_PR_CD_ROOT.resolve())
    except ValueError as exc:
        raise ContractError(
            "CREATE_PR_CD_DEPENDENCY_INVALID",
            "The pinned create-pr-cd entrypoint escapes its package.",
            "dependency",
        ) from exc
    if not entrypoint.is_file():
        raise ContractError(
            "CREATE_PR_CD_DEPENDENCY_INVALID",
            "The pinned create-pr-cd entrypoint is unavailable.",
            "dependency",
        )
    return entrypoint, str(manifest["version"])


def _forward_dependency_event(scope: str, line: str, percent_start: int, percent_end: int) -> None:
    try:
        event = json.loads(line)
    except json.JSONDecodeError:
        return
    if not isinstance(event, dict) or event.get("type") not in {"progress", "warning"}:
        return
    child_percent = event.get("percent")
    mapped_percent = None
    if isinstance(child_percent, (int, float)):
        bounded = max(0.0, min(100.0, float(child_percent)))
        mapped_percent = round(percent_start + ((percent_end - percent_start) * bounded / 100.0))
    phase = str(event.get("phase") or "processing").strip().lower().replace(" ", "_")
    message = str(event.get("message") or f"Generating {scope} entitlement.")
    emit(str(event.get("type")), f"entitlement_{scope.lower()}_{phase}", message, mapped_percent)


def _stop_dependency_process(process: subprocess.Popen[Any], child_cancellation: Path) -> None:
    child_cancellation.parent.mkdir(parents=True, exist_ok=True)
    child_cancellation.touch(exist_ok=True)
    if process.poll() is not None:
        return
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=5)


def run_entitlement_scope(
    envelope: dict[str, Any],
    workspace: Path,
    epms: Path,
    cancellation: Path,
    scope: str,
    percent_start: int,
    percent_end: int,
) -> dict[str, Any]:
    entrypoint, dependency_version = load_create_pr_cd_dependency()
    dependency_workspace = workspace / "temp" / "create-pr-cd" / scope.lower()
    dependency_input = dependency_workspace / "input" / f"epms{epms.suffix.lower()}"
    dependency_output = dependency_workspace / "output"
    dependency_result = dependency_workspace / "result.json"
    child_cancellation = dependency_workspace / "temp" / "cancel.requested"
    if dependency_workspace.exists():
        raise ContractError(
            "CREATE_PR_CD_WORKSPACE_EXISTS",
            f"The isolated {scope} entitlement workspace already exists.",
            "dependency",
            {"scope": scope},
        )
    dependency_input.parent.mkdir(parents=True, exist_ok=True)
    dependency_output.mkdir(parents=True, exist_ok=True)
    child_cancellation.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(epms, dependency_input)
    child_envelope = {
        "schemaVersion": CONTRACT_VERSION,
        "jobId": f"{envelope['jobId']}-entitlement-{scope.lower()}",
        "skill": {"skillId": CREATE_PR_CD_SKILL_ID, "version": dependency_version},
        "parameters": {
            "scope": scope,
            "allSites": True,
            "siteCodes": [],
            "nonProductionUat": False,
        },
        "files": [
            {
                "name": "site_data",
                "path": dependency_input.relative_to(dependency_workspace).as_posix(),
                "originalName": epms.name,
                "size": dependency_input.stat().st_size,
                "sha256": sha256(dependency_input),
            }
        ],
        "paths": {
            "workspace": ".",
            "output": "output",
            "result": "result.json",
            "cancellation": "temp/cancel.requested",
        },
    }
    input_manifest = dependency_workspace / "skill-input.json"
    input_manifest.write_text(json.dumps(child_envelope, ensure_ascii=False, indent=2), encoding="utf-8")
    stdout_log = dependency_workspace / "temp" / "create-pr.stdout.log"
    stderr_log = dependency_workspace / "temp" / "create-pr.stderr.log"
    stream_queue: queue.Queue[tuple[str, str]] = queue.Queue()
    process = subprocess.Popen(
        [sys.executable, str(entrypoint), "--input-manifest", str(input_manifest)],
        cwd=dependency_workspace,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
    )

    def read_stream(name: str, stream: Any, log_path: Path) -> None:
        with log_path.open("w", encoding="utf-8") as log:
            for raw_line in iter(stream.readline, ""):
                log.write(raw_line)
                log.flush()
                stream_queue.put((name, raw_line.rstrip("\r\n")))
        stream.close()

    readers = [
        threading.Thread(target=read_stream, args=("stdout", process.stdout, stdout_log), daemon=True),
        threading.Thread(target=read_stream, args=("stderr", process.stderr, stderr_log), daemon=True),
    ]
    for reader in readers:
        reader.start()
    try:
        while process.poll() is None or any(reader.is_alive() for reader in readers) or not stream_queue.empty():
            if cancellation.exists():
                _stop_dependency_process(process, child_cancellation)
                raise CancelledError()
            try:
                stream_name, line = stream_queue.get(timeout=0.2)
            except queue.Empty:
                continue
            if stream_name == "stdout":
                _forward_dependency_event(scope, line, percent_start, percent_end)
        return_code = process.wait()
    except BaseException:
        _stop_dependency_process(process, child_cancellation)
        raise
    finally:
        for reader in readers:
            reader.join(timeout=2)

    if not dependency_result.is_file():
        raise ContractError(
            "CREATE_PR_CD_RESULT_MISSING",
            f"create-pr-cd did not produce a result for {scope} entitlement.",
            "dependency",
            {"scope": scope, "exitCode": return_code},
        )
    try:
        result = json.loads(dependency_result.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ContractError(
            "CREATE_PR_CD_RESULT_INVALID",
            f"create-pr-cd produced an invalid result for {scope} entitlement.",
            "dependency",
            {"scope": scope, "exitCode": return_code},
        ) from exc
    if return_code != 0 or result.get("status") not in {"succeeded", "succeeded_with_warning"}:
        child_error = result.get("error") or {}
        raise ContractError(
            f"CREATE_PR_CD_{scope}_FAILED",
            f"create-pr-cd could not generate the required {scope} entitlement.",
            "dependency",
            {
                "scope": scope,
                "dependencyCode": child_error.get("code"),
                "dependencyMessage": child_error.get("message"),
                "exitCode": return_code,
            },
        )
    workbooks = []
    for index, item in enumerate(result.get("outputs") or []):
        path = resolve_inside(dependency_workspace, item.get("path"), f"create-pr-cd.{scope}.outputs[{index}].path")
        if path.is_file() and path.suffix.lower() in {".xlsx", ".xlsm"}:
            workbooks.append(path)
    if not workbooks:
        raise ContractError(
            "CREATE_PR_CD_ENTITLEMENT_EMPTY",
            f"create-pr-cd produced no auditable {scope} ECC workbooks.",
            "domain_processing",
            {"scope": scope},
        )
    return {
        "scope": scope,
        "status": result["status"],
        "workbooks": workbooks,
        "warnings": result.get("warnings") or [],
        "metrics": (result.get("summary") or {}).get("metrics") or {},
    }


def run_entitlement_generation(
    envelope: dict[str, Any],
    workspace: Path,
    epms: Path,
    cancellation: Path,
) -> list[dict[str, Any]]:
    results = []
    ranges = ((10, 35), (35, 60))
    for scope, (percent_start, percent_end) in zip(ENTITLEMENT_SCOPES, ranges):
        check_cancel(cancellation)
        emit("progress", f"entitlement_{scope.lower()}", f"Generating mandatory {scope} ECC entitlement.", percent_start)
        results.append(
            run_entitlement_scope(
                envelope,
                workspace,
                epms,
                cancellation,
                scope,
                percent_start,
                percent_end,
            )
        )
        emit("progress", f"entitlement_{scope.lower()}_completed", f"Generated mandatory {scope} ECC entitlement.", percent_end)
    return results


def run(input_manifest: Path) -> int:
    envelope: dict[str, Any] = {}
    result_path = input_manifest.resolve().parent / "result.json"
    try:
        envelope, workspace, output, result_path, cancellation = load_envelope(input_manifest)
        signal.signal(signal.SIGTERM, lambda *_: (_ for _ in ()).throw(CancelledError()))
        check_cancel(cancellation)
        emit("progress", "contract_validation", "Validated the composite audit request.", 5)
        final_po = declared_files(envelope, workspace, "final_po", False)[0]
        epms = declared_files(envelope, workspace, "epms", False, (".xlsx", ".xlsm"))[0]
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
        entitlement_results = run_entitlement_generation(envelope, workspace, epms, cancellation)
        expected_ecc = [
            workbook
            for entitlement in entitlement_results
            for workbook in entitlement["workbooks"]
        ]
        emit("progress", "entitlement_ready", "Loading generated TSS and TI ECC entitlement.", 65)
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
            pr_model=str(CREATE_PR_CD_ROOT / "Info" / "input" / "pr_model.xlsx"),
            epms=str(epms),
            scope_evidence=None,
            scope_evidence_json=None,
            ecc_sheet=audit_final_po.ECC_SHEET_NAME,
            output=str(audit_output),
            summary_json=str(summary_output),
            annotate_ecc_output=bool(parameters.get("annotateEcc", True)),
            annotated_ecc_output_root=str(output / "annotated-ecc"),
            annotated_ecc_timestamp="current",
        )
        audit_final_po.set_cancellation_probe(lambda: check_cancel(cancellation))
        emit("progress", "workbook_read", "Reading Final PO and generated ECC workbooks.", 70)
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
        candidates.append(create_delivery_archive(output.resolve(), candidates).resolve())
        metrics = safe_metrics(summary)
        metrics["entitlement"] = {
            "dependencySkillId": CREATE_PR_CD_SKILL_ID,
            "dependencyVersion": CREATE_PR_CD_VERSION,
            "scopes": [
                {
                    "scope": item["scope"],
                    "status": item["status"],
                    "workbookCount": len(item["workbooks"]),
                    "metrics": item["metrics"],
                }
                for item in entitlement_results
            ],
        }
        abnormal = sum(count for name, count in metrics["classifications"].items() if str(name).lower() != "normal")
        entitlement_warnings = [
            warning
            for item in entitlement_results
            for warning in item["warnings"]
        ]
        status = "succeeded_with_warning" if abnormal or entitlement_warnings else "succeeded"
        warnings = list(entitlement_warnings)
        if abnormal:
            warnings.append({"code": "AUDIT_FINDINGS_PRESENT", "message": "The audit contains findings requiring review.", "details": {"count": abnormal}})
        payload = {
            "schemaVersion": CONTRACT_VERSION,
            "jobId": envelope["jobId"],
            "skillId": SKILL_ID,
            "skillVersion": SKILL_VERSION,
            "status": status,
            "summary": {"message": "TX PR entitlement generation and audit completed.", "metrics": metrics},
            "outputs": [output_item(path, workspace) for path in candidates],
            "warnings": warnings,
            "error": None,
        }
        write_result(result_path, payload)
        emit("progress", "completed", "TX PR entitlement generation and audit completed.", 100)
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
