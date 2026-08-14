# Integrating with `openweb.ui`

This package should live behind an application-owned service boundary. Do not let an LLM choose arbitrary filesystem paths or call backend libraries directly.

## Recommended layering

```text
chat/tool request
→ auth + authorization
→ resolve uploaded attachment from trusted storage
→ allocate per-user/per-request workspace
→ call one domain package
→ accept typed result/refusal
→ register validated output attachment
→ cleanup by host retention policy
```

The Office package remains provider-agnostic. The Open WebUI layer owns conversation state, approval UX, file storage, and model/tool registration.

## Install in the backend image

Pin a release tag in a slim production image without adding a `git` binary:

```text
kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.3.0.zip
```

For an exact commit, replace `refs/tags/v0.3.0` with its 40-character commit SHA. During local development:

```bash
python -m pip install -e /path/to/office
```

## Suggested project placement

The current `Kwentin3/corp-openweb-ui` repository already uses isolated Python sidecars under `services/` with adjacent `openwebui_actions/`. Follow that boundary instead of copying this runtime into Open WebUI itself:

```text
corp-openweb-ui/
  services/
    office-artifacts/
      Dockerfile
      pyproject.toml
      office_artifacts_service/  # FastAPI/application adapter
      openwebui_actions/         # thin Open WebUI action/tool bridge
      tests/
  compose/
    openwebui.compose.yml         # sidecar wiring only
  data/
    office-work/                  # runtime volume, not source control
```

Add the pinned archive dependency to `services/office-artifacts/pyproject.toml`:

```toml
[project]
requires-python = ">=3.11"
dependencies = [
  "kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.3.0.zip",
  "fastapi>=0.115,<1",
  "uvicorn>=0.30,<1",
]
```

Copy or adapt `examples/openwebui_backend/office_service.py` inside the sidecar service layer; do not copy the Office runtime packages into `corp-openweb-ui`. Keep `openwebui_actions/` thin: it should submit an authenticated sidecar request and register the returned attachment, not perform Office mutation itself.

## Request workspace rules

Use a directory controlled by the host, for example:

```text
<root>/<user-id>/<request-id>/
  input/source.docx
  output/result.docx
  internal/docx/
```

Required host checks:

1. Derive user/request IDs from authenticated server state, not model text.
2. Reject traversal, symlinks, device files, and source/output collisions. The Linux/POSIX sidecar adapter performs staging relative to already-open directory descriptors (`dir_fd` + `O_NOFOLLOW`/`O_DIRECTORY`), verifies directory identity after the copy, and deletes the staged file through the original descriptor if a concurrent rename/symlink swap is detected.
3. Copy an admitted upload into the request workspace before processing.
4. Give each request an isolated internal work directory.
5. Keep output names host-generated or basename-only.
6. Register an output only when the domain result is `ok`/`valid`.
7. Return typed refusals unchanged to the orchestrator; do not retry by weakening policy.
8. Delete request data according to the application's retention policy.

## Typical DOCX edit flow

```python
source = service.stage_upload(request_id, trusted_upload_path, ".docx")
tool = service.docx_tool(request_id)
snapshot = tool.inspect(source, view="full")
planned = tool.plan(snapshot, model_request)
if planned["status"] == "refused":
    return planned
output = service.output_path(request_id, "result.docx")
result = tool.apply(source, planned["plan"], output)
```

The source and output paths must differ. `apply` independently revalidates direct or forged plans.

## Typical XLSX flow

Use `summary` or `search` first, then a bounded `region` snapshot. Formula strings are allowed only through explicit formula operations. If a result reports `formula_recalculation: required`, the UI must not present cached values as recalculated results.

## PPTX mode selection

Select the mode in application code:

- existing branded deck/template → `pptx_artifact_tool`;
- new semantic deck → `pptx_ai_composer`.

Do not silently fall back from preservation-first editing to creation-first rendering.

## Model-facing contract

A model should see only:

- compact agent instructions from `packages/<domain>/AGENT_SKILL.md`;
- JSON schemas where provided;
- bounded snapshots returned by `inspect`;
- typed result/refusal envelopes.

Do not expose local paths, raw XML, `openpyxl`, `python-docx`, `python-pptx`, or arbitrary code execution.

## HTTP mapping

Map domain refusals to a successful tool invocation with a structured body, not an internal server crash. For example:

```json
{"status":"refused","reason":"stale_snapshot","details":"source fingerprint changed"}
```

Use HTTP 4xx only for host-layer authentication, ownership, malformed transport, or quota failures. Domain refusals are part of the normal tool protocol.

## Concurrency

Tool instances may share installed code, but not request work directories. Use a separate domain work directory for each request. Run the sidecar workspace under a dedicated OS identity and do not grant unrelated same-UID processes write access to its root. The example adapter closes the demonstrated staging race with descriptor-relative creation and post-copy inode verification. Publication is atomic at the file boundary; higher-level attachment registration remains the host's transaction.

## Production checklist

- Pin the package revision.
- Run repository suites in the Open WebUI backend image.
- Configure upload and output limits below or equal to domain limits.
- Add Word/Excel/PowerPoint or LibreOffice open-save-reopen workers if strict compatibility is required.
- Pin LibreOffice in the sidecar image before enabling `office-witness`; the Python distribution does not install it.
- Add Excel recalculation before showing computed formula values.
- Track output hashes and domain validation results in application audit logs.
