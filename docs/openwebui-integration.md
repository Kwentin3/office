# Integrating with Open WebUI

This package belongs behind an application-owned service boundary. Do not let an LLM choose arbitrary filesystem paths or call the Office runtime libraries directly.

The integration decision below was checked against the live official Open WebUI documentation on 2026-08-16. The Office-specific layering is this project's architecture; upstream documentation is cited separately so that recommendation is not confused with an Open WebUI guarantee.

**Implementation status:** this repository ships the Office domain packages and the path-safe `OfficeService` example. It does not yet ship the FastAPI/OpenAPI transport endpoints or the optional native Open WebUI Tool/Action described below. Those remain downstream integration work and require pinned-version contract tests before production use.

## Extension choice

Use an **external OpenAPI tool server** as the default boundary for Office work. Add an optional **thin native Workspace Tool or Action** only when the integration needs Open WebUI-local file handling, native `HTMLResponse` conversion, or bidirectional confirmation/input calls. The external sidecar can send the supported one-way UI events through Open WebUI's authenticated event endpoint; that alone does not justify an in-process bridge.

| Open WebUI mechanism | Decision for this project | Evidence boundary |
|---|---|---|
| External OpenAPI tool server | Default Office service boundary | Open WebUI documents external HTTP through OpenAPI/MCP, automatic endpoint discovery, and recommends an external service for hard isolation, independent scaling, or conflicting dependencies. |
| Native Workspace Tool/Action | Optional host bridge only | Native Tools can receive attached-file metadata, return inline Rich UI, and make bidirectional UI calls, but execute Python inside the Open WebUI process. The bridge must not render or mutate Office files. |
| Pipelines | Do not use for new integration work | Open WebUI marks Pipelines as legacy and directs external HTTP services to OpenAPI or MCP tool servers. |
| Native MCP | Not the default for this bounded REST-shaped service | Native MCP starts at Open WebUI v0.6.31, is admin-only, has a wider trust boundary, and the upstream guide says OpenAPI remains preferred for most deployments. |

OpenAPI is the simpler contract here: the Office service already has bounded request/response operations, typed refusals, ordinary HTTP authentication needs, and no requirement for MCP sampling, elicitation, or persistent protocol sessions. This is an architecture choice, not a claim that MCP is unsupported.

### Upstream stability boundary

Open WebUI's official Tool Development guide lists `__files__`, `__event_emitter__`, and `__event_call__` as optional native-tool arguments. However, the detailed Reserved Arguments tutorial is explicitly marked as a community contribution not supported by the Open WebUI team. Therefore:

- pin and integration-test the Open WebUI version used in production;
- depend only on the minimum documented argument names and event payloads required by the bridge;
- do not make Open WebUI's internal upload directory layout part of the Office API;
- do not assume `__files__[...]["file"]["path"]` is always a local path: the upstream tutorial says it may instead be a bucket URL for S3, GCS, or Azure storage;
- stage admitted bytes into the sidecar's request workspace through a host-controlled adapter rather than passing an internal Open WebUI path to the model or treating it as a stable cross-version contract.

## Recommended layering

```text
Open WebUI chat
├─ external OpenAPI operations ───────────────┐
└─ optional thin Workspace Tool/Action        │
   ├─ resolve/register Open WebUI File objects│
   ├─ emit status/files/embeds                │
   └─ call the same authenticated OpenAPI API │
                                               ▼
Office sidecar (FastAPI/OpenAPI)
├─ authn/authz context supplied by the host
├─ per-user/per-request workspace
├─ transport DTO ↔ closed domain request/result
└─ exactly one document-domain package call
                                               ▼
kwentin-office domains
├─ DOCX
├─ XLSX
├─ preservation-first PPTX
└─ creation-first PPTX composer
```

The Open WebUI layer owns conversation state, user identity, approval UX, registered files, and tool configuration. The Office sidecar owns isolation, staging, quotas, typed transport mapping, and retention. Each document package remains the only owner of its semantic model, validation, review packet, mutation, and final export.

There must be one source of truth per creation request:

```text
validated semantic model
→ structural preview + revision-bound ReviewPacket
→ chat revision of the complete model
→ explicit approval of that revision
→ final Office artifact from the same validated revision
```

Preview HTML, SVG, PNG, Open WebUI embeds, and local paths are evidence or delivery handles. None becomes editable document state.

## Open WebUI delivery contract

The package and example sidecar return `display_artifacts` as absolute paths **inside the sidecar request workspace**. That field is a host-neutral adapter result, not an Open WebUI UI protocol and not proof that the browser can read the path.

Keep host transports separate:

- **Hermes WebUI:** a registered/local artifact may be rendered with Hermes-native `MEDIA:` syntax.
- **Open WebUI:** register the output as an Open WebUI File object and attach it through the documented file-event surface, or let a native Tool/Action return a persisted Rich UI embed.

For Open WebUI review UX:

- `status` events are suitable for live progress;
- `files` / `chat:message:files` attach Open WebUI File objects;
- the short `embeds` event name is documented as persisted to the database; use `replace: true` when one revision should replace an earlier widget on the same message;
- a native Tool/Action may return `HTMLResponse` with `Content-Disposition: inline` for a sandboxed Rich UI iframe; the upstream guide describes Rich UI embeds as persistent in chat history;
- `__event_call__` can request confirmation or input but depends on a live browser response and can report a disconnected client. It must not be the sole durable approval record.

Persist approval in application-owned conversation/orchestration state as at least:

```json
{
  "request_id": "server-issued",
  "revision": "sha256-bound-review-revision",
  "approved_by": "authenticated-user-id",
  "approved_at": "server-time",
  "review_contract": "host-owned-reference"
}
```

Before final export, re-read that record and reject if the approved `revision` does not match the current validated semantic model. The browser event is UX; the revision-bound record is authorization evidence.

Open WebUI's checked documentation states that external OpenAPI tool results are complete rather than token-streamed. The Events guide documents authenticated one-way events from external OpenAPI/MCP tools, while confirmation/input remains native-Python-only because it needs bidirectional WebSocket communication. The MCP guide separately says the Open WebUI UI event system is native-Python-only; because those two upstream statements are not fully reconciled, pin Open WebUI and contract-test the external REST event route before depending on it. Keep any fallback native bridge thin rather than moving Office execution into it.

## Install in the sidecar image

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
      office_artifacts_service/  # FastAPI/OpenAPI transport adapter
      tests/
  openwebui_actions/
    office_artifacts.py          # optional thin native Tool/Action bridge
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

Copy or adapt `examples/openwebui_backend/office_service.py` inside the sidecar service layer; do not copy the Office runtime packages into `corp-openweb-ui`. Keep `openwebui_actions/` thin: it may normalize Open WebUI file metadata, call the authenticated sidecar API, and register/attach the returned artifact. It must not parse Office XML, compile layouts, mutate files, or weaken a typed refusal.

### OpenAPI registration

For a shared production deployment, register the sidecar as an **admin-managed Global OpenAPI Tool Server**. Open WebUI's menu labels and locations can change between versions, so verify the exact registration screen against the pinned Open WebUI version. Open WebUI documents Global Tool Server requests as originating from the Open WebUI backend, so the backend network must reach the sidecar URL. Scope access to the intended users/groups and authenticate the sidecar at the network/application boundary.

Global tools are hidden by default and must be enabled for the chat through the **+ tool/integration menu**; administrator registration alone does not make a tool active for every chat.

User Tool Servers are browser-originated, visible only to the registering user, and require the Direct Tool Servers permission for non-admins. Their exact Settings location must likewise be verified against the pinned Open WebUI version. They are useful for local development but are not the recommended multi-tenant production topology for this service.

Expose only closed, bounded operations. Give each operation a stable descriptive name, narrow typed request/response schema, server-enforced byte/work limits, and idempotency semantics where applicable. Open WebUI's checked parser documentation describes discovery of standard OpenAPI 3.x HTTP operations; it does **not** establish every internal Open WebUI field or filesystem path as a public compatibility contract.

## Request workspace rules

Use a directory controlled by the sidecar, for example:

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

Open WebUI chat is the only editing surface in this MVP. Do not build a second editor in the preview pane.

```text
chat turn
→ orchestrator emits a complete revised semantic model
→ sidecar validates the format-specific model
→ domain atomically publishes its closed ReviewPacket + structural preview
→ host registers/displays an Open WebUI File object or persisted Rich UI embed
→ next chat turn
→ after revision-bound approval, domain exports from the same model
```

DOCX and XLSX use separate format-owned HTML preview backends:

```python
docx_review = service.docx_chat_review(request_id, revised_docx_model)
xlsx_review = service.xlsx_chat_review(request_id, revised_xlsx_model)
```

DOCX returns one bounded HTML document preview. XLSX returns one bounded HTML grid per visible sheet, up to its documented review limit; hidden sheets are not displayed and are reported in diagnostics. Both HTML formats escape model content, contain no script or external resources, and expose stable semantic addresses (`block_id`, or sheet name + cell coordinate). They are structural previews, not Word pagination or Excel application rendering. XLSX preview shows formulas as text and never claims recalculation.

The result contains `revision`, `review_contract`, `display_artifacts`, and `interaction: chat_only`. The bridge must read/register those sidecar-owned HTML files and deliver them through the Open WebUI mechanisms described above; it must not expose the absolute sidecar path to the model or browser.

## Chat-only PPTX review loop

```text
chat turn
→ orchestrator emits a complete revised DeckSpec
→ sidecar validates DeckSpec
→ composer atomically publishes ReviewPacket + SVG/PNG
→ host registers/displays PNG File objects or a persisted review embed
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
- `display_artifacts` — absolute per-slide PNG paths inside the request workspace;
- `interaction: chat_only` — explicit MVP boundary.

Register the PNGs as Open WebUI File objects or render a persisted native embed. Do not add PDF.js, Gotenberg, a custom Office editor, direct-edit events, or UI-to-DeckSpec patching to this MVP. The adapter refuses model-selected asset paths unless the authenticated host passes them in `allowed_asset_paths`; the composer then independently verifies every approved file against its declared SHA-256 before publishing review evidence. The structural PNG/SVG review uses labeled placeholders for image nodes and states that limitation in `review.json`; final PPTX export consumes the authenticated raster snapshots.

Ownership remains strict:

- Open WebUI: conversation, identity, approval UX, file registration, and attachment/embed delivery;
- Office sidecar: authenticated request mapping, workspace isolation, staging, quotas, and retention;
- PPTX composer: closed DeckSpec, trusted compilation, preview, and ReviewPacket;
- preview: read-only evidence, never mutation;
- native renderer: final PPTX, never prompt interpretation.

## Model-facing contract

A model should see only:

- compact agent instructions from `packages/<domain>/AGENT_SKILL.md`;
- JSON schemas where provided;
- bounded snapshots returned by `inspect`;
- typed result/refusal envelopes;
- host-owned opaque file/review identifiers when needed, never raw storage paths.

Do not expose local paths, raw XML, `openpyxl`, `python-docx`, `python-pptx`, or arbitrary code execution.

## HTTP mapping

Map domain refusals to a successful tool invocation with a structured body, not an internal server crash. For example:

```json
{"status":"refused","reason":"stale_snapshot","details":"source fingerprint changed"}
```

Use HTTP 4xx only for host-layer authentication, ownership, malformed transport, or quota failures. Domain refusals are part of the normal tool protocol.

## Concurrency

Tool instances may share installed code, but not request work directories. Use a separate domain work directory for each request. Run the sidecar workspace under a dedicated OS identity and do not grant unrelated same-UID processes write access to its root. The example adapter closes the demonstrated staging race with descriptor-relative creation and post-copy inode verification. Publication is atomic at the file boundary; higher-level Open WebUI file registration remains the host's transaction.

## Production checklist

- Pin the Office package revision and Open WebUI version.
- Run repository suites and host-bridge contract tests in the production images.
- Verify OpenAPI discovery and every native reserved argument/event used by the bridge against that pinned Open WebUI version.
- Configure upload and output limits below or equal to domain limits.
- Keep Office processing out of the Open WebUI process.
- Persist approval as an authenticated revision-bound record; do not rely only on a live browser event.
- Add Word/Excel/PowerPoint or LibreOffice open-save-reopen workers if strict compatibility is required.
- Pin LibreOffice in the sidecar image before enabling `office-witness`; the Python distribution does not install it.
- Add Excel recalculation before showing computed formula values.
- Track output hashes and domain validation results in application audit logs.
- Exercise multi-user isolation, reconnect, restart, and retention tests before production sign-off.

## Official Open WebUI references

Checked 2026-08-16:

- [Extensibility overview](https://docs.openwebui.com/features/extensibility/)
- [Pipelines — legacy guidance](https://docs.openwebui.com/features/extensibility/pipelines/)
- [OpenAPI Tool Servers](https://docs.openwebui.com/features/extensibility/plugin/tools/openapi-servers/)
- [Connecting an OpenAPI server to Open WebUI](https://docs.openwebui.com/features/extensibility/plugin/tools/openapi-servers/open-webui/)
- [Native Tool Development](https://docs.openwebui.com/features/extensibility/plugin/tools/development/)
- [Reserved Arguments tutorial](https://docs.openwebui.com/features/extensibility/plugin/development/reserved-args/) — community-contributed and explicitly unsupported by the Open WebUI team
- [Events](https://docs.openwebui.com/features/extensibility/plugin/development/events/)
- [Rich UI Embedding](https://docs.openwebui.com/features/extensibility/plugin/development/rich-ui/)
- [Native MCP support](https://docs.openwebui.com/features/extensibility/mcp/)
