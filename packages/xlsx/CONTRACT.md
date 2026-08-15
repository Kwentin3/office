# XLSX Artifact Tool — Frozen Bounded MVP Contract

Status: contract frozen before runtime implementation.

## 1. Goal

Build a separate XLSX-first tool that is convenient for an LLM agent to create, inspect, safely edit, transform, and validate Excel workbooks without exposing raw OOXML, `openpyxl` objects, arbitrary Python mutation, or model-generated library code.

Public lifecycle:

```text
create → inspect → plan → apply → validate
```

The reusable cross-format idea is transaction discipline, not a common Office object model.

## 2. User pains addressed

1. An agent should not write a new `openpyxl` script for every workbook.
2. A model should not guess sheet/cell coordinates from prose alone.
3. Existing formulas, styles, merged regions, dimensions, freeze panes, filters, validations, hyperlinks, comments, and unrelated sheets must not silently change.
4. All operations must pass preflight before any workbook mutation.
5. The source must remain unchanged; `source == output` is forbidden.
6. Stale, malformed, ambiguous, conflicting, unsupported, unsafe, or excessive plans must produce typed refusals.
7. Formula editing must be explicit; cached-value recalculation must never be fabricated.
8. The result must be built as a private candidate, validated, and atomically published.
9. The host must report exact semantic changes and collateral package changes.

## 3. Agent-facing representation

`inspect` returns a closed JSON snapshot with:

- `source_sha256`, `snapshot_sha256`, workbook metadata;
- ordered sheets with `sheet_id`, title, state, used range, dimensions, freeze panes, filter/table/merge summaries;
- sparse non-empty cells with transaction-scoped `tx_*` IDs;
- for each cell: sheet, coordinate, row, column, scalar value or formula, data type, number format, style fingerprint, merge membership;
- bounded row records for a requested rectangular region;
- warnings for unsupported/high-risk workbook features;
- no raw OOXML and no library objects.

Transaction IDs are derived from the current source/snapshot and never written into the workbook.

Views:

- `summary`: workbook/sheet inventory and bounded previews;
- `region`: one exact sheet and rectangular range;
- `search`: bounded matches by displayed scalar/formula text, with row context.

A full unbounded workbook dump is not part of the public contract.

## 4. Public requests

All top-level envelopes are closed.

```json
{"mode":"create","model":{},"output":"...xlsx"}
{"mode":"inspect","source":"...xlsx","view":"summary|region|search","sheet":"required for region","range":"required for region","query":"required for search"}
{"mode":"plan","snapshot":{},"request":{"operations":[],"transforms":[]}}
{"mode":"apply","source":"...xlsx","plan":{},"output":"...xlsx"}
{"mode":"validate","source":"...xlsx","before":"optional ...xlsx"}
```

Typed outcomes:

```json
{"status":"ok", "...":"..."}
{"status":"refused","reason":"unsupported_capability|ambiguous_target|stale_snapshot|validation_failure|unsafe_plan|conflict","details":"short"}
```

## 5. Create scope

A structured workbook model supports:

- ordered worksheets with unique valid names;
- scalar cells: null, string, boolean, finite number, ISO date/datetime;
- formulas as explicit formula objects, not magic strings;
- bounded cell styles: named semantic style `header|currency|percent|date|integer|text|normal`;
- column widths, row heights, freeze panes and one auto-filter range;
- merged ranges only when explicitly declared and non-overlapping.

All creation references use canonical uppercase A1 notation and Excel bounds. Cell and freeze-pane references are single cells (`A1` itself is not a meaningful freeze pane); auto-filter references are one cell or an ordered rectangle; column-width keys are canonical columns `A..XFD`; row-height keys are canonical positive row numbers; merged ranges are ordered, bounded, non-overlapping rectangles. Aliases that an Office library would normalize are refused so preview and final create consume the exact same validated model.

Create does not support charts, pivots, macros, images, external links, conditional formatting, data validation, comments, threaded comments, slicers, PowerQuery, embedded objects, digital signatures, or arbitrary style dictionaries.

Before final create, `render_xlsx_preview(model, output_dir)` may atomically publish a closed `xlsx_chat_review` packet plus one escaped, script-free, bounded HTML grid per visible sheet (maximum 20). Hidden and over-limit sheets are reported, formulas are displayed as text and never evaluated, and each artifact digest binds its published bytes. This is structural review evidence for chat-only iteration, not an Excel application render or editable workbook state.

## 6. Exact edit primitives

The MVP supports exactly these six deterministic primitives:

1. `set_cell_value`
   - target: exact cell `tx_*` ID;
   - new value: bounded scalar or null;
   - refuses formula targets unless `expected_kind:"value"` matches.

2. `set_cell_formula`
   - target: exact cell `tx_*` ID;
   - formula supplied without external workbook references;
   - cached result is reported as unknown/stale until recalculated by an application.

3. `clear_cell`
   - clears value/formula only; preserves style unless explicitly unsupported by the target state.

4. `set_cell_style`
   - applies one semantic style from the allowlist; no arbitrary style object.

5. `append_rows`
   - appends typed rows to an exact inspected table-like rectangular region;
   - refuses merged targets and inconsistent widths;
   - copies formulas/styles only according to an explicit `copy_from_row_id` contract.

6. `reorder_rows`
   - exact permutation of a contiguous inspected region;
   - row IDs are snapshot-bound;
   - formulas are moved with row contents; formula translation/rewrite is not performed implicitly.

No operation uses a bare sheet name + coordinate as sufficient authority at apply time. Human-readable coordinates remain in the snapshot for usability, but apply resolves transaction IDs against the same source fingerprint.

## 7. Bounded transforms

Transforms generate exact primitive operations during `plan`; apply never executes arbitrary computation.

- `fill_missing`: fill null/blank targets in an inspected region;
- `bulk_replace`: exact scalar replacement in bounded targets;
- `sort_rows`: stable sort by one or more inspected key columns, emitting exact `reorder_rows`;
- `compute_column`: arithmetic `sum|difference|product|ratio` from explicit source cell IDs into exact targets;
- `table_totals`: bounded row totals plus optional grand total.

Division by zero, non-numeric inputs, non-finite values, ambiguous headers, or missing targets return typed refusals.

## 8. Formula policy

- Formulas are preserved as formulas (`data_only=False`).
- The tool does not claim to calculate Excel formulas.
- External workbook references, DDE-like links, unsupported dynamic data connections, and newly introduced unsafe formulas are refused.
- Existing formulas outside the intended target must remain semantically and package-wise preserved.
- Formula edits mark calculation state as requiring application recalculation when supported by the package.
- `validate` reports formula text and preservation; it never invents cached values.

## 9. Complete preflight and conflicts

Before mutation, runtime rechecks the whole plan:

- closed envelope and operation schemas;
- source/snapshot/plan hashes;
- operation count and region/cell budgets;
- every target ID and expected old value/kind;
- merged-cell restrictions;
- duplicate writes to one cell;
- overlap between cell edits and row reorder/append;
- exact row permutation and rectangular cardinality;
- formula safety;
- output path differs from source;
- unsupported workbook features according to policy.

Direct forged calls to `apply` receive the same checks as normal `plan` output.

## 10. Transaction and validation

```text
admit source package
→ inspect current source
→ complete preflight
→ create exclusive private candidate
→ mutate candidate
→ semantic inspect of candidate
→ package/member collateral diff
→ atomic os.replace(candidate, output)
```

Mandatory validation:

- ZIP/OPC safety: duplicates, traversal, absolute/backslash names, symlink/nonregular members, member count/size/ratio budgets;
- XML parse and required XLSX parts/relationships/content types;
- workbook opens with the deterministic XLSX library;
- target-specific semantic postconditions for every primitive;
- expected formula/value/style/row-order outcomes;
- source hash unchanged;
- unexpected changed package members reported and blocking outside an operation-derived allowlist;
- sheet order, sheet names, hidden state, dimensions, merges, panes, filters, tables, formulas, styles, names, validations, comments/hyperlinks and workbook properties compared as applicable.

Application open/save/reopen, formula recalculation, rendering, print layout and visual fidelity are separate HOLD gates.

## 11. Cardinality budgets

Initial hard limits, enforced in schemas, `plan`, and forged `apply`:

- worksheets: 256;
- create cells: 250,000;
- inspect returned cells: 20,000;
- operations/intents: 1,000;
- transform input/target cells: 50,000;
- appended rows: 10,000;
- columns per row: 1,000;
- text/formula length per cell: 32,767 characters;
- ZIP/package budgets: explicit runtime constants and tests.

Exceeding a limit returns `unsafe_plan` or `validation_failure` before mutation.

## 12. Explicit refusal boundary

The bounded MVP refuses mutation of:

- `.xls`, `.xlsb`, `.ods`, password-encrypted workbooks;
- VBA/macros (`.xlsm`) and digital signatures;
- charts, pivot tables/caches, slicers, images/drawings, embedded objects;
- external workbook links, PowerQuery/data connections;
- named-range creation/editing;
- conditional formatting and data validation edits;
- arbitrary merge/unmerge;
- arbitrary insert/delete rows or columns;
- formula translation, recalculation, or cached-value claims;
- arbitrary style dictionaries or theme editing;
- generic range DSL and model-generated Python.

Existing unsupported features may be admitted read-only only if preservation can be demonstrated; otherwise edit requests refuse before mutation.

## 13. Evidence required for bounded PASS

- frozen 30-task benchmark written before implementation;
- RED→GREEN unit/integration tests for every primitive and refusal class;
- create/inspect/plan/apply/validate CLI/API parity;
- exact source preservation and private candidate publication tests;
- formula/style/merge/pane/filter/sheet-order collateral probes;
- malformed/forged/stale/conflicting/adversarial package tests;
- fixed-public LLM-usability suite using only public Skill/schemas/snapshots;
- independent exact-tree review;
- deterministic package, fresh extraction, manifest, checksum.

## 14. Claims

A bounded PASS may claim deterministic structural XLSX creation/editing in the tested subset.

It may not claim:

- Microsoft Excel or LibreOffice compatibility;
- formula-engine correctness or recalculated cached values;
- pixel/print fidelity;
- hidden-input generalization;
- production or hostile multi-tenant readiness.

## v0.2 bounded additions

`inspect(..., view="inventory")` returns a format-specific risk inventory. External links, connections, OLE, macros, and signatures block mutation. `fill_template(source, values, output, strict=True)` supports exact `{{token}}` replacement only in bounded non-formula string cells and compiles to existing `set_cell_value` operations; it adds no new mutation primitive.
