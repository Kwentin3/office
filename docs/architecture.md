# Architecture

## Decision

The repository shares transaction policy, not a universal Office model.

```text
shared principles ≠ shared domain runtime
```

Every domain owns its vocabulary, contract, mutation/compiler layer, validator, and refusal boundary.

## Repository structure

```text
packages/
  docx/           document/story/paragraph/table/cell domain
  xlsx/           workbook/sheet/region/row/cell/formula domain
  pptx-editor/    template/slide/slot/table-cell domain
  pptx-composer/  DeckSpec → SceneSpec → isolated render/preview backends
  application-witness/ private clone → configured application → typed observation

docs/             host and integration documentation
examples/         provider-neutral integration and artifact examples
scripts/          build/test utilities
```

## Common policy sequence

```text
closed public contract
→ trusted deterministic plan/compiler
→ bounded format-specific representation
→ private candidate
→ independent package and semantic validation
→ atomic publication
```

These are duplicated intentionally where semantics differ. A helper should be shared only when its meaning—not merely its code shape—is identical in multiple domains.

## Non-goals

The repository does not provide:

- `OfficeArtifactBase`;
- a universal Office AST;
- a common OOXML mutation runtime;
- a common SceneSpec for DOCX/XLSX/PPTX;
- arbitrary Python or callback plugins;
- raw OOXML or library objects in LLM-facing contracts;
- unrestricted layout coordinates.

## Domain boundaries

### DOCX

Create and preservation-first editing over document stories, paragraphs, tables, rows, and cells. Transaction IDs bind exact semantic targets to a source hash.

The creation path also owns an isolated chat-review backend:

```text
validated DocSpec
→ bounded escaped HTML structural preview
→ closed DOCX ReviewPacket
→ final DOCX from the same DocSpec after approval
```

Stable `document_id` and `block_id` values belong to the semantic model. The HTML is review evidence only and never becomes the source of truth.

### XLSX

Create and preservation-first operations over sheets, bounded rectangular regions, cells, formulas, rows, and managed styles. Formula mode is explicit and recalculation is never fabricated.

The creation path also owns an isolated chat-review backend:

```text
validated SheetSpec/workbook model
→ bounded escaped HTML per visible sheet
→ closed XLSX ReviewPacket
→ final XLSX from the same model after approval
```

Sheet names and cell coordinates remain stable review addresses. Preview displays formula text but never evaluates it or claims current cached values.

### PPTX editor

Template-first, preservation-sensitive edits through placeholders and explicitly named `slot:*` shapes. Apply copies the admitted source into a private immutable snapshot and uses that snapshot for preflight, mutation, collateral comparison, and candidate validation. Existing design and unrelated package parts remain template-owned.

### PPTX composer

Creation-first flow:

```text
DeckSpec
→ trusted recipe compiler
→ bounded SceneSpec
→ native PPTX backend / structural preview backend
→ validator
```

`SceneSpec` is private to this creation-first PPTX domain. The preview backend also publishes a closed `ReviewPacket` for chat orchestration. It contains only the DeckSpec revision, fidelity boundary, bounded diagnostics, and hash-bound SVG/PNG artifact names. Existing review generations are exchanged atomically on Linux; cleanup after that commit point is best effort. Image nodes are labeled placeholders in preview and are rendered from authenticated byte snapshots only by the native PPTX backend. The packet contains no prompt, conversation history, UI event, mutation command, or renderer coordinates.

### Chat/WebUI adapter

The provider-specific adapter is outside every document runtime. It may allocate request workspaces, invoke exactly one domain API, and register returned artifacts. For review publication it retains descriptor-relative authority through the complete domain call and rechecks every bound request ancestor before returning. It must not parse prompts, compile layouts, edit Office files, or translate direct UI manipulation into document mutations. The MVP interaction is chat-only; every visual revision begins with a complete validated domain model.

For Hermes WebUI, the verified downstream pattern is a native five-tool plugin (`start`, `preview`, `approve`, `export`, `status`) plus a separate pinned `kwentin-office` Python process. That is dependency/process isolation, not an OS sandbox. The host owns chat/turn identity, direct-user approval provenance, durable revision state, private workspaces, and `MEDIA:` delivery. The Office packages return host-neutral paths and revisions; they neither emit `MEDIA:` nor know Hermes session identifiers. The tested adapter currently refuses external PPTX assets until a host-approved attachment channel exists. The downstream plugin is not currently distributed by this repository; see `docs/hermes-integration.md` for the exact status and verification boundary.

For Open WebUI, the default deployment is an external OpenAPI Office sidecar plus an optional thin native Workspace Tool/Action for host-local file handling, native Rich UI conversion, or bidirectional confirmation/input. Supported one-way events may be sent from the sidecar through Open WebUI's authenticated event endpoint, subject to pinned-version contract tests. Pipelines are not part of the architecture. Native MCP is not required for the bounded REST-shaped Office contract. Absolute `display_artifacts` paths are sidecar-local adapter outputs, not browser URLs or an Open WebUI transport API.

### Application witness

The witness is an OS/application boundary, not a shared Office model. It knows only the closed artifact type, file suffix, subprocess policy, and output validation. It receives no mutation plan and cannot publish a document. Rich Inspect and strict template compilers remain inside their format packages.

## Host responsibilities

The package validates artifacts, not user identity or tenant ownership. A host such as Hermes WebUI or Open WebUI must provide authentication, authorization, revision-bound approval persistence, request isolation, upload policy, quotas, cleanup, and native artifact delivery. The host adapter must stage admitted bytes without making host-internal storage paths part of the document-domain contract.
