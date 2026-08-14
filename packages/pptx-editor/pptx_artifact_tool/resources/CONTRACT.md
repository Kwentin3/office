# PPTX Artifact Tool — Frozen Template-First MVP Contract

Status: frozen before runtime implementation.

## Goal

Provide an LLM agent a small, safe, provider-agnostic API for repeatable corporate PPTX rendering and bounded edits:

```text
reference/template PPTX → inspect/extract normalized JSON → plan → apply → validate
```

The template owns slide count, layouts, masters, geometry, decorative shapes, theme, fonts and visual style. The agent fills explicit native placeholders or named slots. It never receives raw OOXML, `python-pptx` objects, or arbitrary mutation code.

## Public lifecycle

```text
create(template, model, output)
inspect(source, view)
plan(snapshot, request)
apply(source, plan, output)
validate(source, before?)
```

`create` means render an existing template deck; it does not mean free-form layout generation.

## Stable template slots

Editable slots are:

1. native title/body/subtitle placeholders; or
2. text/table shapes explicitly named `slot:<globally-unique-key>`.

Decorative text, charts, SmartArt, groups and unnamed arbitrary shapes are read-only.

The normalized schema exposes:

- template/source SHA-256;
- ordered slide transaction IDs;
- slide title and layout name;
- slot key, kind, role, shape name and current content;
- table cell transaction IDs;
- bounded warnings and unsupported-feature inventory.

`tx_*` IDs bind mutations to exact source content. Stable `slot_key` values are only for `create(template, model, output)` and are resolved to transaction IDs before apply.

## Inspect views

- `summary`: slide order, titles, layout names, slot inventory, warnings;
- `slide`: one exact slide with editable slots and table cells;
- `search`: bounded text matches with slide/shape context.

No unbounded raw XML dump.

## Exact primitives

1. `set_slot_text`
   - exact text-slot transaction ID;
   - single or multiline Unicode string;
   - preserves shape geometry and first-run/paragraph formatting as available.

2. `clear_slot_text`
   - exact text-slot ID;
   - preserves the shape and its formatting container.

3. `set_table_cell_text`
   - exact table-cell transaction ID;
   - preserves table geometry and cell formatting container.

4. `reorder_slides`
   - exact permutation of all current slide transaction IDs;
   - changes presentation slide order only.

No bare slide index + shape name is sufficient authority at apply time.

## Template render (`create`)

Model:

```json
{
  "slots": {
    "cover.title": "Quarterly review",
    "cover.subtitle": "Q3 2026",
    "summary.body": "Line 1\nLine 2"
  },
  "table_cells": {
    "results.metric.r1c1": "Revenue"
  },
  "slide_order": ["cover", "summary", "results"]
}
```

All keys must resolve exactly once. Unknown, duplicate or missing required keys refuse before mutation. The model compiles to the same primitives as manual plan.

## Transaction and validation

```text
admit source package
→ inspect exact source
→ closed-contract preflight
→ exclusive private candidate
→ package-preserving mutation of only operation-derived XML parts
→ semantic inspect candidate
→ package collateral diff
→ atomic publication
```

Validation requires:

- ZIP safety: duplicate/traversal/backslash/absolute/nonregular members and bounded size/count/ratio;
- required presentation parts and parseable XML;
- source/snapshot/plan hashes;
- private immutable source snapshot for every apply transaction;
- exact old target content and kind;
- duplicate/overlapping target refusal;
- expected slot/cell text or slide order after mutation;
- source unchanged;
- changed package parts limited to affected `ppt/slides/slideN.xml` and, for reorder, `ppt/presentation.xml`;
- all unrelated package members byte-preserved.

## Cardinality budgets

- slides: 500;
- slots returned: 5,000;
- table cells returned: 20,000;
- operations: 1,000;
- text per slot/cell: 32,767 characters;
- search results: 100;
- package members: 10,000;
- total uncompressed package: 128 MiB;
- single member: 32 MiB.

## Explicit refusal boundary

The MVP refuses mutation of:

- `.ppt`, `.pps`, `.odp`, encrypted or macro-enabled decks;
- adding/deleting/cloning slides;
- changing slide layouts, masters, themes, fonts or geometry;
- arbitrary shape movement/resizing/creation/deletion;
- grouped shapes, SmartArt, charts, equations, OLE and embedded objects;
- image/media replacement;
- animations/transitions;
- hyperlinks/actions;
- speaker notes and comments;
- table row/column insertion/deletion or merge changes;
- autofit/font-size synthesis;
- model-generated Python or raw OOXML.

Existing unsupported parts are admitted read-only and byte-preserved. An operation targeting them refuses.

## Evidence required for bounded PASS

- frozen 30-task benchmark before implementation;
- RED→GREEN tests for all four primitives and refusal classes;
- CLI/API parity;
- template extraction and render by stable keys;
- stale/forged/conflict/adversarial package tests;
- exact package preservation checks;
- real PPTX dogfood and user application-open confirmation;
- fixed-input LLM usability suite;
- independent exact-tree review;
- fresh-extraction release package and checksum.

## Claims and HOLD

A bounded PASS may claim deterministic package-preserving template-slot text/table rendering and slide reordering in the tested subset.

HOLD:

- PowerPoint/LibreOffice compatibility until opened there;
- pixel/visual fidelity until rendered and compared;
- text overflow, font substitution and pagination-like layout behavior;
- free-form slide design;
- image/chart/SmartArt/media mutation;
- production/hostile multi-tenant readiness;
- hidden-input generalization.

## v0.2 bounded additions

`inspect(..., view="inventory")` reports charts, SmartArt, media, notes, animations, hyperlinks, comments, OLE, macros, and signatures. OLE, macros, and signatures block mutation. `fill_template(source, values, output, strict=True)` compiles exact `{{token}}` replacements only for managed `slot:*` text/table targets into existing primitives. Unmanaged tokens and notes refuse.
