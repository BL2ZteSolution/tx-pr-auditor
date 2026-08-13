# tx-pr-auditor

Standalone composite PR-audit skill for AI Worker Platform.

Its public contract accepts a Final PO workbook and an EPMS workbook. The
Python entrypoint runs the pinned `create-pr-cd` dependency for mandatory TSS
and TI entitlement, then audits Final PO against those generated ECC files.
The platform does not sequence or interpret either engine.

Audit scope is resolved from column B of the pinned create-pr-cd PR Model
before description heuristics. Exact EPMS PR references authorize TSS, TI,
and Planning audit-only entitlement; TX Integrated evidence authorizes
Operation Back Office. This reconstruction is audit-only and does not alter
create-pr-cd duplicate prevention or generate production PR output.

Clone recursively, then run:

```text
git clone --recurse-submodules <repository>
python src/main.py --input-manifest <workspace>/skill-input.json
```

Use `scripts/audit_final_po.py` directly only for focused Final PO-versus-ECC
development and testing.
