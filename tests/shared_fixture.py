import os
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def approval_context():
    platform = Path(os.environ["TEST_PLATFORM_ROOT"]) if os.environ.get("TEST_PLATFORM_ROOT") else ROOT.parents[1]
    mw = Path(os.environ["CREATE_PR_CD_ROOT"]) if os.environ.get("CREATE_PR_CD_ROOT") else ROOT.parent / "create-pr-cd"
    approval = json.loads((platform / "backend/src/skills/approvedSkills.json").read_text(encoding="utf-8"))["create-pr-cd"]
    files = sorted({p.relative_to(mw).as_posix() for pattern in approval["include"] for p in mw.glob(pattern) if p.is_file() and "__pycache__" not in p.parts})
    return {"skillId": "create-pr-cd", "version": approval["approvedVersion"], "packageSha256": approval["approvedPackageSha256"], "root": str(mw), "runtimeFiles": files}
