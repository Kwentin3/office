# Application Witness

`office_application_witness` is an isolated operational adapter. It does not share an Office document model with DOCX, XLSX, or PPTX backends and never performs publication.

## Boundary

```text
validated source path
→ descriptor-opened private clone
→ isolated LibreOffice profile and HOME
→ fixed argv subprocess (shell=False)
→ bounded output validation
→ typed observation report
→ unconditional workspace cleanup
```

The source must be a bounded regular non-symlink file and is copied through an already-open descriptor. The configured executable and work directory must be absolute host-controlled paths. Application stdout/stderr are discarded rather than accumulated. The output directory is descriptor-pinned before process launch; the result is opened relative to that descriptor, and directory identity plus source/output size and modification metadata are rechecked after hashing. The private output is opened with `O_NOFOLLOW`, bounded, hashed and deleted before a success report is returned. Cleanup failure is a typed refusal. The report contains hashes and statuses, not a round-tripped document.

The adapter controls paths, environment, argv and process-group lifetime on every leader outcome, but it is **not an OS sandbox**. `process_isolation` is therefore `trusted_executable_not_sandboxed`: the host must pin a trusted LibreOffice binary/container. An arbitrary executable with the service account's permissions could write outside the private workspace.

## API

```python
from office_application_witness import ApplicationWitness

witness = ApplicationWitness(
    "/srv/office-work/request-123/witness",
    executable="/usr/bin/soffice",
    runtime_identity={
        "application_version": "LibreOffice 26.2.5.2",
        "image_digest": "sha256:<64 lowercase hex characters>",
    },
)
report = witness.observe(
    "/srv/office-work/request-123/output/result.docx",
    "docx",
    timeout_seconds=60,
)
```

Supported observations:

- DOCX/PPTX: conversion of a private clone to a bounded PDF and PDF signature validation;
- XLSX: conversion of a private clone to a separate XLSX round-trip and ZIP validation.

The generic XLSX operation asks LibreOffice to recalculate during round-trip, but the generic adapter does not semantically verify cached formula values. The report therefore says `requested_not_semantically_verified`. The separate `scripts/application_gates/run_xlsx_recalculation.py` gate performs a private fixed-argv round-trip and then compares only declared cached values, with explicit type/tolerance and unsupported-formula cases.

Success and refusal results are closed `TypedDict` contracts with `schema_version: 1`. Runtime identity is host-supplied, closed to `application_version` and a full `sha256:` image digest, bounded and validated before process launch. When absent, both fields are `not_observed`. JSON stdin cannot choose or override this identity, the executable, or the work directory; unbounded application output is never captured.

## Claims intentionally not made

- `repair_dialog` is `not_observable_headless`;
- Microsoft Office equivalence is `not_claimed`;
- a LibreOffice PASS is not pixel-level fidelity evidence;
- fake-soffice CI tests validate subprocess lifecycle, private-path handling and report semantics, not OS sandboxing or real application compatibility.

## Runtime gate

The package does not install LibreOffice. The host sidecar must provide and pin the absolute executable path. Repository CI is hermetic and uses fake executables. `.github/workflows/libreoffice-compat.yml` is a manual-only gate requiring a self-hosted runner labeled `office-libreoffice-pinned` and host-controlled runtime identity variables. It generates non-sensitive fixtures, records exact commit/tree plus runtime identity, runs DOCX/XLSX/PPTX observations, and runs the declared XLSX cached-value gate. No schedule is enabled until a real pinned host produces a stable evidence series.
