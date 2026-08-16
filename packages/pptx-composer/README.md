# AI PPTX Composer — bounded vertical slice

A contract-first, native-PPTX tracer bullet for the larger chat-driven presentation workflow.

## What works now

```text
closed semantic DeckSpec JSON
→ trusted recipe compiler
→ closed bounded SceneSpec (text/shape/image/chart + budgets)
→ isolated fast-preview and native-PPTX backends
→ managed catalog of components/archetypes
→ 7 semantic archetypes
→ native editable text/shapes/charts
→ hash-bound raster assets
→ sanitized SVG source + hash-bound PNG fallback
→ private candidate
→ structural/semantic/bounded-geometry validation
→ atomic publication
```

Archetypes:

- `cover`
- `comparison`
- `chart_with_takeaway`
- `process`
- `timeline`
- `decision_matrix`
- `kpi_grid`

The renderer creates editable PowerPoint text and shapes. Charts are native chart parts with an embedded XLSX workbook.

## Architecture boundary

This project is deliberately **not** the whole product:

- conversation, brief collection and approvals belong to the orchestrator;
- source extraction/provenance belongs to the evidence domain;
- image generation/search belongs to the asset domain;
- this project currently implements the closed DeckSpec contract, native composition subset and validation subset;
- preservation-sensitive editing remains in the sibling `pptx_artifact_tool` package.

Critical invariants are marked in code with `STICKY INVARIANT`, `STICKY BOUNDARY`, `STICKY LIMIT`, or `STICKY CLAIM BOUNDARY`. These comments explain boundaries that future refactors must not erase.

## Run

```bash
office-pptx-compose <<'JSON'
{"action":"render","spec":"examples/pptx-composer/managed-library.deck.json","output":"deck.pptx"}
JSON

office-pptx-compose <<'JSON'
{"action":"validate","spec":"examples/pptx-composer/managed-library.deck.json","source":"deck.pptx"}
JSON
```

## Test

```bash
python -m unittest discover -s packages/pptx-composer/tests -t packages/pptx-composer -v
```

## Fast feedback loop

The public JSON-lines surface now includes:

- `catalog` — inspect managed components/archetypes and their lifecycle state;
- `preview` — compile the same DeckSpec into an atomically published chat-review bundle;
- `render` — create the native editable PPTX;
- `validate` — run the available PPTX gates.

The MVP interaction is intentionally chat-only:

```text
Hermes chat revises DeckSpec
→ preview publishes ReviewPacket + per-slide SVG/PNG
→ Hermes WebUI displays the PNGs with its native MEDIA support
→ user and LLM review
→ the next chat turn revises DeckSpec again
```

`DeckSpec` is the only source of truth. Preview artifacts are read-only evidence and never mutate the model. The renderer does not receive prompts, conversation history, UI events, or arbitrary coordinates.

Each preview directory contains:

- `review.json` — closed `pptx_chat_review` V1 contract;
- `manifest.json` — backend-oriented structural manifest;
- one SVG and one 1280×720 PNG per selected slide.

The ReviewPacket is bound to the exact validated DeckSpec plus compiled preview selection/variants by a canonical SHA-256 `revision`. Before publication, preview admits every declared asset through the same boundary as native rendering: it opens paths non-blocking without following symlinks, immediately rejects non-regular files, snapshots and hash-authenticates bounded files, fully decodes PNG/JPEG bytes with declared-format checks, sanitizes SVG, and independently validates its PNG fallback. Native rendering consumes the admitted immutable bytes rather than reopening caller-controlled paths. The packet declares `interaction: chat_only`, carries bounded diagnostics and hash-binds every SVG/PNG. The CLI response returns `review_contract` and absolute `display_artifacts`; a host adapter should register or display those files through its own native file/media surface rather than building another viewer.

Example preview request:

```json
{
  "action": "preview",
  "spec": "examples/pptx-composer/managed-library.deck.json",
  "output": "managed-library-preview"
}
```

`preview` and `render` also accept closed optional controls for local iteration:

```json
{
  "slide_ids": ["s_compare"],
  "variants": {"s_compare": "compact"}
}
```

Variant names come from the managed catalog. Unknown targets, malformed arrays/objects and unsupported variants are refused rather than ignored.

The preview manifest deliberately reports:

```text
structural_preview_not_powerpoint_render
```

It is optimized for immediate conversational review. Its bounded wrapper reports likely text overflow as node-level diagnostics, but this remains an approximation based on the bundled preview font—not evidence of exact PowerPoint metrics, chart fidelity or repair-free application opening.

The managed catalog lives in `pptx_ai_composer/library.py`. Archetypes are named compositions of bounded components; the catalog validates component references, variants, status (`experimental`, `stable`, `deprecated`) and preview fidelity. The compiler reads allowed variants from this catalog so the list has one source of truth. The catalog contains no executable Python or LLM-provided coordinates.

## Honest validation boundary

Current validator executes:

- ZIP admission and duplicate-member check;
- package size/member budgets;
- XML parseability;
- required package parts;
- parser reopen;
- slide count;
- expected text;
- presence of native chart;
- shape bounding boxes against slide canvas.

Not executed in this environment:

- raster rendering;
- PowerPoint-accurate text overflow detection (the structural preview only reports bounded proxy diagnostics);
- pixel-level collisions and visual hierarchy;
- font substitution;
- Microsoft PowerPoint open/repair/save round trip.

A clean report is therefore `valid_with_unexecuted_gates`, never an unconditional production `VALID`.

## SVG policy

`python-pptx 1.0.2` in this environment cannot directly ingest SVG. V1 therefore:

1. hash-binds the original SVG;
2. parses it with network/entity resolution disabled;
3. refuses scripts, foreign objects, active attributes and external references;
4. hash-binds a supplied PNG fallback;
5. embeds the fallback and does **not** claim native-vector output.

Direct SVG embedding can be added only after a renderer spike proves it in the target PowerPoint versions.

## Deferred by design

- prompt/document → brief and outline planning;
- image provider adapters and stock search;
- arbitrary layouts or raw coordinates;
- tables and generic diagrams;
- speaker notes/citations as native note parts;
- render-to-images and visual critic;
- slide replacement/diff API;
- reference-slide retrieval;
- animations, transitions, video, collaboration and analytics.

These are domain milestones, not fields to add opportunistically to one mega-class.
