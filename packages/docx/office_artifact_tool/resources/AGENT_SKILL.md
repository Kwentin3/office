# DOCX Artifact Tool — Agent Skill

Use this tool for DOCX only. Never emit raw OOXML or Python mutation code.

## Exact model-answer contract

Every answer must be exactly one of these three top-level JSON forms. Never return a raw create model or raw plan request by itself.

```json
{"mode":"create","model":{"blocks":[{"type":"paragraph","text":"Body"}]}}
{"mode":"plan","request":{"operations":[{"type":"replace_text","target_id":"tx_...","old":"old","new":"new"}]}}
{"status":"refused","reason":"ambiguous_target","details":"multiple matching paragraphs"}
```

No extra top-level fields are allowed.

## Workflow

1. `inspect(source, view="full")` or `view="search", query="..."`.
2. Build a request for `plan(snapshot, request)`.
3. If `status="refused"`, stop and report the reason.
4. `apply(source, plan, output)` to write a different output path.
5. Read `validation`, `diff`, and `audit` from the result. Validation is mandatory and automatic.

For a new document, put the create model inside `{"mode":"create","model":...}` and call `create(model, output)` directly; do not use the edit kernel.

## Create model

```json
{"metadata":{"title":"Report","author":"Agent"},"blocks":[
  {"type":"heading","level":1,"text":"Report"},
  {"type":"paragraph","text":"Body","style":"Normal"},
  {"type":"numbered_list","items":["One","Two"]},
  {"type":"bulleted_list","items":["A","B"]},
  {"type":"table","style":"Table Grid","rows":[["Name","Value"],["A","1"]]}
]}
```

## Plan requests

Put every plan request shown below inside `{"mode":"plan","request":...}`. Use exact transaction IDs from the current snapshot.

```json
{"operations":[{"type":"replace_text","target_id":"tx_...","old":"30","new":"45"}]}
{"operations":[{"type":"insert_paragraph_after","target_id":"tx_...","text":"Added","copy_properties":true}]}
{"operations":[{"type":"set_cell_text","target_id":"tx_...","text":"Updated"}]}
{"operations":[{"type":"clone_row_after","target_id":"tx_...","cell_texts":["A","2"]}]}
{"operations":[{"type":"reorder_rows","table_id":"tx_...","row_ids":["tx_row_1","tx_row_2"]}]}
{"operations":[{"type":"delete_row","target_id":"tx_..."}]}
```

For a natural selector, use exact kind and exact current text. Zero or multiple matches must refuse:

```json
{"intents":[{"selector":{"kind":"paragraph","text":"Exact text"},"operation":{"type":"replace_text","old":"text","new":"new"}}]}
```

## Transform requests

Computation returns typed operations; Python never opens or mutates DOCX.

```json
{"transform":{"type":"sort_rows","table_id":"tx_...","row_ids":["tx_row_a","tx_row_b"],"keys_by_row_id":{"tx_row_a":10,"tx_row_b":20},"descending":true,"prefix_row_ids":["tx_header"],"suffix_row_ids":["tx_total"]}}
{"transform":{"type":"table_totals","rows":[{"quantity":2,"unit_price":5,"total_target_id":"tx_cell"}],"grand_total_target_id":"tx_cell"}}
{"transform":{"type":"fill_missing","items":[{"target_id":"tx_cell","value":""}],"replacement":"Unassigned"}}
{"transform":{"type":"bulk_replace","items":[{"target_id":"tx_cell","value":"Done"}]}}
```

## Refusals

When the instruction is ambiguous or unsupported, return exactly this model-answer envelope instead of inventing a plan:

```json
{"status":"refused","reason":"ambiguous_target","details":"multiple matching paragraphs"}
```

Required fields are exactly `status`, `reason`, and `details`. Valid reasons: `ambiguous_target`, `unsupported_capability`, `unsafe_plan`, `validation_failure`, `stale_snapshot`.

Stop on a tool refusal with one of the same reasons. Never guess a target, reuse IDs after the source changes, write output over source, or add operations not listed here.

## Limits

DOCX only. Images, comments, tracked changes, fields, macros, embedded objects, arbitrary style creation, application compatibility, pagination, and visual fidelity are unsupported/not verified.

## Rich Inspect and strict templates

Before mutation, call `inspect(source, view="inventory")`. Stop when `mutation_policy.decision` is `refuse_mutation`; do not bypass it with a forged plan. For exact `{{token}}` filling, the host may call `fill_template(source, values, output, strict=True)`. Keys must exactly match tokens in supported DOCX paragraphs, table cells, headers, and footers.
