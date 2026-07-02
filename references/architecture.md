# TX PR Auditor Architecture

## Purpose

TX PR Auditor validates submitted TX PR/PO records against Final PO, EPMS, and PR Model data. It classifies submitted records as:

- Normal
- Abnormal - Invalid PO
- Abnormal - Wrong PO
- Abnormal - Duplicate PO

The business rules are governed by `TX_PR_Auditor_Business_Logic_Specification_v1.0`. This reference defines only the technical architecture.

## Product Boundary

`create-pr-cd` is the execution framework. TX PR Auditor is an independent skill mounted inside that framework.

Framework responsibilities:

- User interaction
- File selection
- File validation
- Workflow control
- Passing file paths to the skill
- Receiving audit results
- Saving output

TX PR Auditor responsibilities:

- Reading source workbooks
- Normalizing data
- Matching EPMS
- Generating business facts
- Resolving PR Model mappings
- Generating expected items
- Executing audit decisions
- Resolving duplicates
- Writing the audit report

The framework must never contain TX business rules.

## Input Model

The framework passes explicit paths to three files:

- Final PO workbook
- EPMS workbook
- PR Model workbook

Rules:

- The skill never searches for files.
- Final PO uses worksheet `条目明细`.
- EPMS uses worksheet `data`.
- PR Model is one workbook per execution.
- The user supplies the correct PR Model workbook.

Expected workspace shape:

```text
Workspace/
├── Final PO.xlsx
├── EPMS.xlsx
├── PR Model.xlsx
└── Output/
```

## Pipeline Model

Use a centralized batch pipeline:

```text
Workbook Reader
Field Mapper
Canonical Builder
EPMS Matcher
Business Fact Builder
PR Model Resolver
Expected Item Generator
Audit Engine
Duplicate Resolver
Report Writer
```

Process entire datasets stage by stage. Do not process row by row as the primary architecture, because duplicate calculation requires visibility of the entire Final PO snapshot.

Only the pipeline orchestrator controls execution. Modules never call each other directly.

## Immutable Datasets

Each stage emits a new typed dataset and does not mutate previous outputs:

```text
Raw Dataset
Canonical Dataset
Business Dataset
Expected Dataset
Audit Dataset
Duplicate Dataset
Output Dataset
```

Use strongly typed collections instead of generic row collections after mapping.

Example canonical dataset:

```text
Canonical Dataset
├── Final PO Records
├── EPMS Records
└── Metadata
```

Example business dataset:

```text
Business Dataset
├── Final PO Business Records
├── EPMS Business Records
├── Business Facts
└── Metadata
```

Example expected dataset:

```text
Expected Dataset
├── Expected Items
├── Expected Quantities
└── Metadata
```

Example audit dataset:

```text
Audit Dataset
├── Audit Results
└── Metadata
```

Example duplicate dataset:

```text
Duplicate Dataset
├── Final Results
└── Metadata
```

## Stage Responsibilities

Workbook Reader:

- Read workbook contents.
- Perform no business processing.
- Emit Raw Dataset.

Field Mapper:

- Convert Excel columns into canonical fields.
- Example fields: DU, Item Code, Quantity, Dispatch Date, Settlement Quantity, Dispatch Status.
- Perform no audit logic.

Canonical Builder:

- Create standardized business records.
- Make downstream processing independent from Excel layout.

EPMS Matcher:

- Match Final PO records to EPMS records.
- Use DU as the primary matching key.
- Emit matched business records.

Business Fact Builder:

- Derive business facts from EPMS.
- Examples: Region, Planning Subcontractor, Integration End Date, Antenna Size, Configuration, TX SOW.

PR Model Resolver:

- Use the supplied PR Model workbook.
- Determine applicable worksheet and business mapping.
- Use one PR Model workbook per execution.

Expected Item Generator:

- Generate expected entitlement.
- Examples: Expected Item Code, Expected Quantity, Expected Scope, Expected Subcontractor.
- Do not compare submitted records here.

Audit Engine:

- Compare submitted records with expected entitlement.
- Follow the governing business logic specification.
- Emit audit decisions only.

Duplicate Resolver:

- Process only valid claims.
- Consume expected quantities in Dispatch Date, Request Number, Dispatch Order Number, PO Line Number order.
- Ensure Wrong PO and Invalid PO do not consume quantity.

Report Writer:

- Generate `PR_Audit_Result.xlsx`.
- Add audit result columns and evidence fields.

## Frozen Decisions for Version 1.0

- Workspace-based execution is confirmed.
- Separate Final PO, EPMS, and PR Model files are confirmed.
- A single PR Model workbook is confirmed.
- Framework-passed explicit file paths are confirmed.
- Modular pipeline architecture is confirmed.
- Batch processing is confirmed.
- Immutable datasets are confirmed.
- Central pipeline orchestration is confirmed.
- Dataset ownership by stage is confirmed.
- Strongly typed datasets are confirmed.
