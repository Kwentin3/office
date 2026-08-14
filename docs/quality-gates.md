# Quality gates

## Current bounded evidence

| Domain | Local suite | Independent bounded review | User-open smoke test |
|---|---:|---|---|
| PPTX composer | 66/66 | APPROVE | PASS |
| DOCX | 36/36 | APPROVE | PASS |
| XLSX | 37/37 | APPROVE | PASS |
| PPTX editor | 19/19 | post-fix independent re-review pending | PASS |
| Open WebUI adapter | 5/5 | symlink/exclusive/cleanup/directory-race path-safety tests | N/A |

The repository CI reruns the packaged source rather than trusting these historical counts.

## What a local PASS means

- closed contract and typed refusals;
- source/output collision protection;
- source fingerprint and plan binding;
- private candidate and validation before publication;
- bounded package admission;
- tested semantic postconditions;
- source remains unchanged in tested flows.

## Unexecuted production gates

- automated Microsoft Office or LibreOffice open-save-reopen;
- repair-dialog detection;
- pixel/render fidelity and font substitution;
- Word pagination;
- Excel formula recalculation and compatibility;
- hostile multi-tenant deployment controls.

Therefore the honest status is:

```text
bounded implementation: APPROVED
production compatibility: HOLD until host-specific application gates
```

A user's successful manual opening is valuable smoke-test evidence, but not a substitute for automated application round trips across supported versions.
