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
kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.4.0.zip
```

For an exact commit, replace `refs/tags/v0.4.0` with its 40-character commit SHA. During local development:

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
  "kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.4.0.zip",
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
2. Reject traversal, symlinks, device files, and source/output collisions. The Linux/POSIX sidecar adapter performs staging relative to already-open directory descriptors (`dir_fd` + `O_NOFOLLOW`/`O_DIRECTORY`), verifies directory identity after the copy, and deletes the staged file through the original descriptor if a concurrent rename/symlink swap is detected. Chat-review publication likewise keeps its domain directory descriptor open through the complete renderer call and refuses if any bound ancestor changes.
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

## Chat-only creation review loops

The MVP uses Hermes/Open WebUI as the only editing surface. Do not build a second editor in the preview pane.

```text
chat turn
→ orchestrator emits a complete revised semantic model
→ sidecar validates the format-specific model
→ domain atomically publishes its closed ReviewPacket + structural preview
→ host registers/displays native HTML or PNG media
→ next chat turn
→ after approval, domain exports the Office file from the same model
```

DOCX and XLSX use separate format-owned HTML preview backends:

```python
docx_review = service.docx_chat_review(request_id, revised_docx_model)
xlsx_review = service.xlsx_chat_review(request_id, revised_xlsx_model)
```

DOCX returns one bounded HTML document preview. XLSX returns one bounded HTML grid per visible sheet, up to its documented review limit; hidden sheets are not displayed and are reported in diagnostics. Both HTML formats escape model content, contain no script or external resources, and expose stable semantic addresses (`block_id`, or sheet name + cell coordinate). They are structural previews, not Word pagination or Excel application rendering. XLSX preview shows formulas as text and never claims recalculation.

The result contains `revision`, `review_contract`, `display_artifacts`, and `interaction: chat_only`. Register each HTML path as a normal attachment or emit its registered/local path through Hermes WebUI native `MEDIA:` handling.

## Chat-only PPTX review loop

The MVP uses Hermes/Open WebUI as the only editing surface. Do not build a second editor in the preview pane.

```text
chat turn
→ orchestrator emits a complete revised DeckSpec
→ sidecar validates DeckSpec
→ composer atomically publishes ReviewPacket + SVG/PNG
→ host registers/displays PNG attachments
→ next chat turn
```

Call the example adapter with a server-issued request ID:

```python
review = service.pptx_chat_review(
    request_id,
    revised_deck_spec,
    allowed_asset_paths=tuple(host_approved_asset_paths),
)
```

The result contains:

- `revision` — canonical SHA-256 of the exact validated DeckSpec plus compiled preview selection/variants;
- `review_contract` — absolute path to closed `review.json`;
- `display_artifacts` — absolute per-slide PNG paths;
- `interaction: chat_only` — explicit MVP boundary.

For Hermes WebUI, register the PNGs as normal attachments or emit the registered/local paths through its native `MEDIA:` response syntax. Do not add PDF.js, Gotenberg, a custom HTML renderer, direct-edit events, or UI-to-DeckSpec patching to this MVP. The adapter refuses model-selected asset paths unless the authenticated host passes them in `allowed_asset_paths`; the composer then independently verifies every approved file against its declared SHA-256 before publishing review evidence. The structural PNG/SVG review uses labeled placeholders for image nodes and states that limitation in `review.json`; final PPTX export consumes the authenticated raster snapshots.

Ownership remains strict:

- Hermes/Open WebUI: conversation, auth, revisions requested by the user, attachment delivery;
- PPTX composer: closed DeckSpec, trusted compilation, preview and ReviewPacket;
- preview: read-only evidence, never mutation;
- native renderer: final PPTX, never prompt interpretation.

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
