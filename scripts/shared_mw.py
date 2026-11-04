"""Fail-closed verification of the one shared, externally approved MW package."""
from __future__ import annotations
import hashlib
import json
import os
import subprocess
from pathlib import Path

ENV = "TX_PR_AUDITOR_MW_DEPENDENCY"
TEXT = {".json", ".md", ".py", ".toml", ".txt", ".yaml", ".yml"}

class SharedDependencyError(ValueError):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code

def resolve_mw_root():
    raw = os.environ.get(ENV)
    if raw:
        try:
            root = Path(json.loads(raw)["root"])
            if not root.is_absolute():
                raise ValueError("absolute root required")
            return root.resolve()
        except (ValueError, KeyError, TypeError):
            raise SharedDependencyError("CREATE_PR_CD_DEPENDENCY_INVALID", "Shared MW approval root is invalid.") from None
    return Path(__file__).resolve().parents[2] / "create-pr-cd"

def verify_shared_mw():
    try:
        approval = json.loads(os.environ[ENV])
        root = resolve_mw_root()
        manifest = json.loads((root / "skill.json").read_text(encoding="utf-8"))
        own = json.loads((Path(__file__).resolve().parents[1] / "skill.json").read_text(encoding="utf-8"))
        declaration = next(d for d in own["dependencies"] if d["skillId"] == "create-pr-cd")
        if manifest.get("skillId") != "create-pr-cd" or approval.get("skillId") != "create-pr-cd" or manifest.get("version") != approval.get("version") or manifest["version"] not in declaration["supportedVersions"]:
            raise SharedDependencyError("CREATE_PR_CD_DEPENDENCY_IDENTITY_MISMATCH", "Shared MW version/identity is incompatible with this Auditor release.")
        files = approval["runtimeFiles"]
        if not isinstance(files, list) or not files or len(files) != len(set(files)):
            raise ValueError("invalid file list")
        required = {"skill.json", "src/main.py", "Info/input/pr_model.xlsx"}
        if not required.issubset(files):
            raise ValueError("incomplete approval")
        digest = hashlib.sha256()
        for name in sorted(files):
            if not isinstance(name, str) or "\\" in name or Path(name).is_absolute():
                raise ValueError("invalid file path")
            path = (root / name).resolve()
            if not path.is_relative_to(root) or path == root:
                raise ValueError("file outside root")
            content = path.read_bytes()
            if path.suffix.lower() in TEXT:
                content = content.replace(b"\r\n", b"\n")
            digest.update((name + "\0" + hashlib.sha256(content).hexdigest() + "\n").encode("utf-8"))
        if digest.hexdigest() != approval["packageSha256"]:
            raise SharedDependencyError("CREATE_PR_CD_DEPENDENCY_FINGERPRINT_MISMATCH", "Shared MW runtime differs from its approved fingerprint.")
        entrypoint = (root / manifest["runtime"]["entrypoint"]).resolve()
        if not entrypoint.is_relative_to(root) or not entrypoint.is_file():
            raise ValueError("invalid entrypoint")
        commit = None
        if (root / ".git").exists():
            read = subprocess.run(["git", "-C", str(root), "rev-parse", "HEAD"], capture_output=True, text=True, timeout=5)
            if read.returncode != 0:
                raise ValueError("commit unavailable")
            commit = read.stdout.strip()
            if approval.get("commit") and commit != approval["commit"]:
                raise SharedDependencyError("CREATE_PR_CD_DEPENDENCY_COMMIT_MISMATCH", "Shared MW checkout differs from its approved commit.")
        return {"skillId": "create-pr-cd", "version": manifest["version"], "packageSha256": digest.hexdigest(), "commit": commit, "root": root, "entrypoint": entrypoint}
    except SharedDependencyError:
        raise
    except (KeyError, TypeError, ValueError, OSError, StopIteration, subprocess.SubprocessError):
        raise SharedDependencyError("CREATE_PR_CD_DEPENDENCY_UNAPPROVED", "Provide a complete trusted shared MW approval through TX_PR_AUDITOR_MW_DEPENDENCY; no nested fallback is used.") from None
