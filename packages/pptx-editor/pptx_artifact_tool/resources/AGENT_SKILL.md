# PPTX Artifact Tool — Agent Skill

Use this tool for **template-first** PPTX work. The template owns design, slide geometry, masters, layouts and media. The agent fills explicit slots; it does not design slides with arbitrary coordinates.

## Workflow

```text
inspect(summary) → inspect(slide/search) → plan → apply → validate
```

For repeatable generation:

```text
template PPTX → inspect slot keys → create(template, JSON model) → validate
```

Editable slots must be native text placeholders/text boxes/tables explicitly named:

```text
slot:<globally-unique-key>
```

Examples: `slot:cover.title`, `slot:summary.body`, `slot:results.table`.

## Inspect

```json
{"action":"inspect","source":"template.pptx","view":"summary"}
{"action":"inspect","source":"template.pptx","view":"slide","slide_id":"tx_..."}
{"action":"inspect","source":"template.pptx","view":"search","query":"Revenue"}
```

Use returned `tx_*` IDs for edits. Never invent IDs or use a bare slide index/shape name as mutation authority.

## Create from template

```json
{
  "action":"create",
  "template":"template.pptx",
  "model":{
    "slots":{
      "cover.title":"Quarterly Review",
      "summary.body":"Revenue up\nCosts stable"
    },
    "table_cells":{
      "results.table.r1c1":"125"
    }
  },
  "output":"deck.pptx"
}
```

Unknown or ambiguous keys refuse.

## Exact operations

```json
{"type":"set_slot_text","target_id":"tx_...","text":"New text","expected_text":"Old text"}
{"type":"clear_slot_text","target_id":"tx_...","expected_text":"Old text"}
{"type":"set_table_cell_text","target_id":"tx_...","text":"125","expected_text":"100"}
{"type":"reorder_slides","slide_ids":["tx_slide_2","tx_slide_1"]}
```

## Safety

- Source and output must differ.
- Text/table edits change only operation-derived slide XML.
- Slide reorder changes only `ppt/presentation.xml`.
- Masters, layouts, themes, media and unrelated slides remain byte-identical.
- Formula-like strings have no special meaning in PPTX.
- Application rendering and text overflow are separate visual gates.

## Refuse

Refuse arbitrary shape geometry, slide cloning/deletion, layout/master/theme editing, images, charts, SmartArt, groups, animations, notes/comments, hyperlinks/actions, table structure changes, raw OOXML and generated Python.

## Rich Inspect and strict templates

Use `inspect(source, view="inventory")` before mutation and stop on `refuse_mutation`. Strict `fill_template(source, values, output)` resolves `{{token}}` only in explicit `slot:*` text shapes and their table cells. Tokens in unmanaged shapes or speaker notes refuse; no editable surface is inferred.
