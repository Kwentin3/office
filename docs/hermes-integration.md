# Integrating with Hermes WebUI

`kwentin-office` is the document-domain runtime. A Hermes plugin is a separate,
thin host adapter that exposes closed tools, owns chat/session authorization,
runs the package in a pinned Python process, and publishes returned artifacts
through Hermes' native media transport. This process/dependency isolation is
not an OS security sandbox.

Do not put document parsing, OOXML mutation, layout compilation, or approval
policy into an LLM prompt or an arbitrary shell command.

## Implementation status

The `v0.4.0` repository and distribution provide the Office domain packages,
closed semantic models, structural preview backends, final exporters, and the
path-safe `OfficeService` example. They do **not** currently ship an installable
Hermes plugin.

A downstream Hermes plugin using the contract below was independently deployed
and tested on 2026-08-16. That evidence validates the integration pattern and
the tested adapter snapshot; it does not turn that downstream plugin into part
of the `kwentin-office` release. Until its source and installer are published in
this repository, there is no copy-and-paste supported Hermes installation
command here.

Keep this distinction explicit:

- **verified package contract:** shipped by `Kwentin3/office`;
- **verified downstream adapter:** deployed outside this repository;
- **public reusable Hermes integration:** still to be packaged and released.

## Recommended boundary

```text
Hermes WebUI chat
→ native Hermes plugin tools
→ host-owned Office session and approval state
→ fixed-argv isolated kwentin-office runtime
→ exactly one format-domain operation
→ typed result or refusal
→ host validates local artifacts and emits MEDIA: handles
```

The plugin is not an Office implementation. It may:

- create an opaque server-owned Office session;
- bind that session to the host-injected Hermes chat identifier;
- pass one complete validated semantic model to preview;
- retain the exact revision and approval state;
- invoke the pinned runtime with fixed argv and a bounded environment;
- validate and expose returned local artifacts through Hermes;
- export only the exact approved revision.

It must not:

- let the model choose arbitrary server paths, executables, or workspaces;
- accept a replacement semantic model during approval or export;
- approve on behalf of the user;
- parse Office XML or duplicate format-domain validation;
- treat preview HTML, SVG, PNG, or `MEDIA:` handles as editable document state;
- claim structural preview is Microsoft Office rendering.

## Native tool contract

The tested adapter exposed one toolset with five closed operations:

| Tool | Responsibility |
|---|---|
| `office_start` | Create an opaque, server-owned DOCX/XLSX/PPTX review session. |
| `office_preview` | Validate one complete semantic model and publish review artifacts; the tested adapter currently refuses every external PPTX asset. |
| `office_approve` | Record direct user approval for one exact displayed revision in a later turn. |
| `office_export` | Export only the approved revision to a validated caller-supplied basename under a server-owned output directory. |
| `office_status` | Read durable review, approval, and export state for the bound session. |

Model-facing schemas should reject unknown fields and expose neither arbitrary
URLs/paths nor a generic command field. If Hermes progressive tool disclosure
is enabled, the functions may be discovered through
`tool_search → tool_describe → tool_call`; the underlying dispatch still goes
through the native Hermes registry and plugin handler.

The installed adapter refuses PPTX `asset_ref` input until Hermes supplies a
host-approved attachment allowlist. Text, tables, charts, and managed layout
remain available, but external images cannot be used through the current five-
tool contract. Do not silently drop those assets or pass model-selected paths.

Hermes WebUI uses the agent platform toolset selected by the deployed Hermes
runtime. In the tested deployment that surface was `cli`; there was no separate
`webui` platform key. Treat that as a checked deployment fact, not a universal
name to hard-code into `kwentin-office`.

## Required review state machine

```text
office_start(format)
→ office_preview(session, complete semantic model)
→ display every returned preview artifact
→ stop the turn
→ user sends the exact approval phrase in a later turn
→ office_approve(session, exact revision, exact phrase)
→ office_export(session, exact revision, safe basename)
```

The host must enforce all of these invariants:

1. The Office session belongs to the current Hermes chat.
2. Preview and approval cannot occur in the same host turn.
3. The approval phrase comes from the direct current user message, not generated
   assistant text or tool output.
4. Approval is bound to the full exact revision, even if the human-facing phrase
   contains only a safe prefix.
5. Every new preview invalidates previous approval.
6. Export accepts neither a semantic model nor an arbitrary output directory.
7. Cross-chat status, preview, approval, and export are refused.

The tested adapter used host-injected chat and turn identifiers and, where the
plugin executor did not pass direct user text, read only the latest active user
message from Hermes' session database in read-only/query-only mode. This is a
specific downstream implementation technique, not a requirement that belongs
inside the Office package. Another host may provide equivalent non-model-
controlled approval provenance through a different stable API.

## Artifact delivery

The package returns `display_artifacts` as absolute paths rooted at the output
location supplied by its caller. In the tested adapter, the host selects a
bound private workspace and verifies every returned path against it. The paths
are adapter results, not browser URLs and not a Hermes protocol by themselves.

For Hermes WebUI the adapter must:

1. verify that every path remains inside the bound private workspace;
2. register or expose the artifact only through Hermes-approved local media
   roots;
3. return/display a `MEDIA:/absolute/path` handle in chat;
4. preserve the semantic revision and artifact digest alongside the review;
5. never confuse `MEDIA:` with Open WebUI File objects or Rich UI events.

DOCX/XLSX HTML is escaped structural evidence. PPTX SVG/PNG is a structural
composition preview. These are fast review surfaces, not Word pagination,
Excel recalculation/layout, PowerPoint rendering, or Microsoft Office fidelity
proof.

## Runtime and storage isolation

Prefer a pinned Office runtime separate from Hermes' own Python environment.
The plugin should launch a fixed interpreter and worker with `shell=False`,
bounded input/output, a minimal fixed environment, and a timeout. Private
session state belongs under an application-owned root with restrictive
permissions; the model receives only opaque session/revision identifiers. This
is dependency and process separation, not an OS sandbox claim.

The host must retain descriptor-relative authority across publication and
cleanup, reject symlinks and directory replacement, and fail closed when the
runtime output is missing, relocated, non-regular, or invalid. Do not make a
workspace path part of the public model-facing contract.

## Verification checklist

A deployment is not proven by package import or plugin registration alone.
Require all of the following:

1. `kwentin-office` is pinned to an exact released version or commit.
2. Plugin Doctor imports the plugin and reports all five registrations.
3. The Office toolset is enabled for the actual WebUI agent surface.
4. Raw and progressively assembled model catalogs can discover the tools.
5. Registry dispatch reaches every handler and returns bounded typed results.
6. A clean new chat executes the native discovery/call path.
7. DOCX, XLSX, and PPTX each complete preview → later approval → export →
   independent reopen.
8. Cross-chat use is refused for status, preview, approval, and export.
9. Same-turn approval, wrong phrase, stale revision, and approval after a new
   preview are refused.
10. Output-directory replacement and symlink probes publish no escaped files.
11. Runtime/render failure removes incomplete output and returns a typed error.
12. Preview is visibly delivered before approval is requested.

Internal downstream evidence from 2026-08-16 records that the tested snapshot
passed Plugin Doctor with 5/5 tools, the
three-format review/export/reopen flow, four cross-chat refusals, all-format
output-directory replacement probes with zero escaped files, and an independent
exact-hash review with no reproducible medium-or-higher issue. A stable public
verification report has not yet been published in this repository.

## Evidence boundary

The verified downstream run did not prove:

- Microsoft Word/Excel/PowerPoint pixel fidelity;
- real Microsoft Office open/save/reopen or repair-dialog behavior;
- Excel formula recalculation correctness;
- a reusable public installer for the Hermes plugin;
- external PPTX assets through the current adapter, which refuses them until a
  host-approved attachment channel exists;
- portability of the tested Linux filesystem publication primitives to every
  operating system;
- production multi-tenant sandboxing or retention policy.

Keep production status bounded by those unexecuted gates. For Open WebUI, use
the separate [Open WebUI integration guide](openwebui-integration.md); its File
objects, Rich UI, event surfaces, and deployment topology are not Hermes
`MEDIA:` contracts.
