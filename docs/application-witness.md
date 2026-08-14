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

The XLSX operation asks LibreOffice to recalculate during round-trip, but the adapter does not semantically verify cached formula values. The report therefore says `requested_not_semantically_verified`.

Success and refusal results are closed `TypedDict` contracts with `schema_version: 1`. The adapter deliberately reports `version: not_observed`; a production image may collect version evidence separately, but unbounded application output is never captured by this adapter.

## Claims intentionally not made

- `repair_dialog` is `not_observable_headless`;
- Microsoft Office equivalence is `not_claimed`;
- a LibreOffice PASS is not pixel-level fidelity evidence;
- fake-soffice CI tests validate subprocess lifecycle, private-path handling and report semantics, not OS sandboxing or real application compatibility.

## Runtime gate

The package does not install LibreOffice. The host sidecar must provide and pin the absolute executable path. The JSON stdin contract cannot choose either the executable or work directory; those are host-controlled required CLI arguments/configuration. This repository's CI is hermetic and uses a fake executable. A deployment that needs application evidence must add a separate job using its actual pinned LibreOffice image and record the image/application version beside the witness report.
