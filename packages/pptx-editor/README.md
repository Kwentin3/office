# PPTX Artifact Tool

Template-first bounded PPTX generation/editing for an LLM agent.

## Core idea

```text
PPTX template (design + geometry + media)
  → inspect into JSON slots
  → plan exact tx-bound edits
  → package-preserving apply
  → semantic/package validate
```

Editable native text boxes/placeholders/tables are explicitly named `slot:<unique-key>` in PowerPoint's Selection Pane.

## Public lifecycle

- `inspect`: `summary`, bounded `slide`, or `search`.
- `plan`: closed operations with expected old text and transaction-bound IDs.
- `apply`: private candidate, apply-side plan validation, exact postconditions, operation-derived package allowlist, atomic publication.
- `validate`: package/parser checks and package-member diff.
- `create`: compile a JSON slot model into the same plan/apply primitives.

## MVP operations

- `set_slot_text`
- `clear_slot_text`
- `set_table_cell_text`
- `reorder_slides`

See `contracts/PPTX_MVP_CONTRACT.md` and `agent/skill/SKILL.md`.

## Run tests

```bash
python -m unittest discover -s packages/pptx-editor/tests -v
```

## CLI example

```bash
printf '%s' '{"action":"inspect","source":"template.pptx","view":"summary"}' \
  | office-pptx-edit
```

The CLI accepts one JSON request on stdin and returns one JSON response on stdout.

## Current evidence boundary

The runtime preserves non-target PPTX package members byte-for-byte and reopens generated output with `python-pptx`. Microsoft PowerPoint/LibreOffice rendering, overflow detection, animations, charts, SmartArt, images, notes/comments and freeform geometry remain outside this bounded MVP.
