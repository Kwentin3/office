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

### XLSX

Create and preservation-first operations over sheets, bounded rectangular regions, cells, formulas, rows, and managed styles. Formula mode is explicit and recalculation is never fabricated.

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

`SceneSpec` is private to this creation-first PPTX domain.

### Application witness

The witness is an OS/application boundary, not a shared Office model. It knows only the closed artifact type, file suffix, subprocess policy, and output validation. It receives no mutation plan and cannot publish a document. Rich Inspect and strict template compilers remain inside their format packages.

## Host responsibilities

The package validates artifacts, not user identity or tenant ownership. A host such as `openweb.ui` must provide authentication, authorization, request isolation, upload policy, quotas, cleanup, and attachment delivery.
