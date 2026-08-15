# XLSX Artifact Tool — Agent Skill

Use only JSON lifecycle calls. Never emit `openpyxl`/Python code, raw OOXML, bare package paths, or additional fields.

## Workflow

```text
existing workbook: inspect(summary/search) → inspect(region) → plan → apply → validate
new workbook: create model → render_xlsx_preview → chat corrections → approval → create → validate
```

`render_xlsx_preview(model, output_dir)` publishes bounded HTML for visible sheets and a closed ReviewPacket. It displays formula text without evaluating it and is not an Excel application render. The host displays the returned artifacts through native media; preview never becomes editable state.

The host must always validate writes. Never overwrite source; output must be a different `.xlsx` path.

## Inspect

```json
{"action":"inspect","source":"book.xlsx","view":"summary"}
{"action":"inspect","source":"book.xlsx","view":"search","query":"Pending"}
{"action":"inspect","source":"book.xlsx","view":"region","sheet":"Data","range":"A1:D20"}
```

Use `summary` to learn sheet names/layout, `search` to localize content, then `region` to get exact `tx_*` cell/row/region IDs. IDs belong only to the current source fingerprint; never invent or reuse them after the source changes.

## Create

Cells are explicit objects. A formula is never a magic value string.

```json
{"action":"create","model":{"sheets":[{"name":"Data","cells":{"A1":{"value":"Item","style":"header"},"B1":{"value":"Amount","style":"header"},"A2":{"value":"A"},"B2":{"value":10,"style":"currency"},"C2":{"formula":"=B2*2"}},"freeze_panes":"A2","auto_filter":"A1:C2","column_widths":{"A":24,"B":14},"row_heights":{"1":28},"merged_ranges":[]}]},"output":"new.xlsx"}
```

Styles: `normal|header|currency|percent|date|integer|text`.

Creation references are canonical uppercase A1 notation: cells and `freeze_panes` are single cells, `auto_filter` is a single cell or rectangular range, column-width keys are `A..XFD`, row-height keys are canonical positive row numbers, and merged ranges are ordered non-overlapping rectangles. Invalid or normalized aliases are refused before preview or workbook publication.

## Plan — exact primitives

```json
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"set_cell_value","target_id":"tx_...","value":5,"expected_kind":"value"}]}}
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"set_cell_formula","target_id":"tx_...","formula":"=B2*C2","expected_kind":"value"}]}}
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"clear_cell","target_id":"tx_...","expected_kind":"value"}]}}
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"set_cell_style","target_id":"tx_...","style":"currency"}]}}
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"append_rows","region_id":"tx_...","copy_from_row_id":"tx_...","rows":[["B",20],["C",30]]}]}}
{"action":"plan","snapshot":{},"request":{"operations":[{"type":"reorder_rows","region_id":"tx_...","row_ids":["tx_row_2","tx_row_1"]}]}}
```

`append_rows` requires exact region width and explicit template row. `reorder_rows` requires an exact permutation of all rows in the inspected contiguous region.

## Plan — bounded transforms

Transforms compile to the same exact primitives; they do not mutate independently.

```json
{"action":"plan","snapshot":{},"request":{"transforms":[{"type":"fill_missing","target_ids":["tx_..."],"value":0}]}}
{"action":"plan","snapshot":{},"request":{"transforms":[{"type":"table_totals","rows":[{"quantity_id":"tx_...","unit_price_id":"tx_...","target_id":"tx_..."}]}]}}
{"action":"plan","snapshot":{},"request":{"transforms":[{"type":"sort_rows","region_id":"tx_...","keys_by_row_id":{"tx_row_1":"B","tx_row_2":"A"},"descending":false}]}}
```

## Apply / validate

```json
{"action":"apply","source":"book.xlsx","plan":{},"output":"edited.xlsx"}
{"action":"validate","source":"edited.xlsx","before":"book.xlsx"}
```

Formula text is preserved, but this tool does not calculate formulas. A formula edit reports `formula_recalculation:"required"`; cached values remain unknown until Excel/LibreOffice recalculates.

## Typed refusal

On unsupported, stale, malformed, ambiguous, conflicting, unsafe, or excessive requests, return the tool refusal unchanged:

```json
{"status":"refused","reason":"unsupported_capability|ambiguous_target|stale_snapshot|validation_failure|unsafe_plan|conflict","details":"short"}
```

## Unsupported boundary

Refuse `.xls/.xlsb/.ods/.xlsm`, VBA/macros/signatures, charts/pivots/slicers/images/embedded objects, external workbook links, PowerQuery/data connections, named-range edits, conditional-format/data-validation edits, arbitrary merge/unmerge, row/column insertion/deletion, formula translation/recalculation, arbitrary style dictionaries, generic range DSL, and generated Python.

## Rich Inspect and strict templates

Use `inspect(source, view="inventory")` before mutation and stop on `refuse_mutation`. Strict `fill_template(source, values, output)` resolves `{{token}}` only in bounded non-formula string cells and compiles each replacement to `set_cell_value`. Missing, unknown, formula-hosted, malformed, or unresolved tokens refuse.
