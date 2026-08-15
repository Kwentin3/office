# Domain artifact contracts

Status: V1 boundary map. `DeckSpec` is executable; upstream contracts are frozen interface targets, not falsely claimed runtime features.

## Sticky architecture rule

```text
ConversationOrchestrator
  → ContextPack
  → DeckBrief
  → Outline
  → DeckSpec
  → SceneSpec
  → AssetRecord[]
  → PPTX candidate
  → ValidationReport
```

Each arrow is a versioned, closed, bounded JSON artifact. A downstream domain may propose a revision but must not silently rewrite an approved upstream artifact.

No contract may contain:

- OOXML;
- `python-pptx`/PptxGenJS objects;
- executable code or callbacks;
- renderer coordinates in an LLM-facing artifact;
- provider-specific generation fields outside an asset provider adapter.

## C1. `ContextPack` — evidence domain

**Runtime status:** `not_implemented`; existing DOCX/XLSX/web tools are intended adapters.

Required fields:

```json
{
  "contract_version": "1.0",
  "context_id": "context_demo",
  "sources": [
    {"source_id": "metrics", "label": "Metrics.xlsx", "sha256": "..."}
  ],
  "claims": [
    {
      "claim_id": "claim_01",
      "text": "...",
      "support": [{"source_id": "metrics", "location": "Data!D12:D27"}],
      "status": "supported"
    }
  ],
  "datasets": [],
  "conflicts": [],
  "unknowns": []
}
```

Invariants:

- every claim support points to a known source;
- evidence text and source data remain distinct;
- contradiction is a first-class result, not silently resolved;
- source identity is hash-bound;
- no layout or narrative fields.

## C2. `DeckBrief` — planning domain

**Runtime status:** `not_implemented`; conversation/orchestration milestone.

```json
{
  "contract_version": "1.0",
  "brief_id": "brief_demo",
  "goal": "Approve a four-week pilot",
  "audience": ["Operations director", "IT director"],
  "duration_minutes": 12,
  "tone": "evidence-first",
  "language": "ru",
  "brand_profile_id": "brand_default",
  "evidence_policy": "cite_material_claims",
  "image_policy": "user_or_generated_with_provenance",
  "approval": {"status": "approved", "revision": 2}
}
```

Invariants:

- questions belong to the orchestrator, answers belong here;
- only material uncertainties block progress;
- approved revision is immutable;
- no slide order or geometry.

## C3. `Outline` — planning domain

**Runtime status:** `not_implemented`; planning milestone.

```json
{
  "contract_version": "1.0",
  "outline_id": "outline_demo",
  "brief_id": "brief_demo",
  "context_id": "context_demo",
  "slides": [
    {
      "slide_id": "problem",
      "role": "problem",
      "message": "Manual hand-offs consume measurable time",
      "claim_ids": ["claim_01"],
      "visual_intent": "chart_with_takeaway"
    }
  ],
  "approval": {"status": "approved", "revision": 3}
}
```

Invariants:

- one key message per slide;
- stable IDs, never execution by list position;
- every material message maps to claims or is explicitly marked as framing/opinion;
- split/merge/reorder create a new revision;
- no asset paths or coordinates.

## C4. `DeckSpec` — composition boundary

**Runtime status:** `implemented` in `pptx_ai_composer/contracts.py`.

Executable V1 supports exactly:

- `cover`;
- `comparison`;
- `chart_with_takeaway`;
- `process`;
- `timeline`;
- `decision_matrix`;
- `kpi_grid`.

The schema is closed recursively and bounded to 20 slides. Source references, asset references and chart dimensions are validated before rendering. Charts require an explicit `data_source_id` that resolves to a declared source and is cited by the slide; numeric equality is not treated as evidence of placeholder data.

Why `DeckSpec` currently combines approved slide content with brand data: it is the smallest useful executable boundary for the tracer bullet. Do not split it into more runtime objects until a demonstrated workflow needs independent revisions of those objects.

## C5. `SceneSpec` — trusted composition/runtime boundary

**Runtime status:** implemented in `pptx_ai_composer/scene_contract.py`; compiled by `compiler.py`.

`SceneSpec` is not LLM-facing. It contains only bounded `text`, `shape`, `image`, and `chart` nodes with validated canvas boxes, stable IDs and per-deck/per-slide budgets. Archetype recipes and coordinates are owned by trusted compiler code. The preview and native-PPTX backends consume this same artifact.

## C6. `AssetRecord` — asset domain

**Runtime status:** implemented by the `assets` section of `DeckSpec` and the shared `asset_admission.py` boundary.

Raster record:

```json
{
  "asset_id": "hero",
  "kind": "png",
  "path": "/frozen/hero.png",
  "sha256": "...",
  "alt_text": "..."
}
```

SVG record additionally requires:

```json
{
  "fallback_path": "/frozen/hero.png",
  "fallback_sha256": "..."
}
```

Current asset claim boundary:

- both preview and native rendering snapshot each declared asset once through a non-blocking, no-symlink descriptor, immediately reject non-regular files, and authenticate the snapshot against `sha256` before publication;
- each source or fallback is at most 20 MiB; decoded raster dimensions are at most `8192 × 8192` and 40 million pixels;
- PNG/JPEG bytes must fully decode and match the declared kind;
- SVG source is hash-bound and checked for active content/external references through the same admission boundary;
- SVG fallback is independently authenticated, fully decoded and required to be PNG;
- admission returns immutable byte snapshots, and native rendering embeds those snapshots without reopening caller-controlled paths;
- `python-pptx` cannot ingest SVG directly in this environment;
- renderer embeds the authenticated PNG fallback snapshot;
- vector editability is `not_supported`, not implied.

## C7. `ValidationReport` — validation domain

**Runtime status:** implemented in `pptx_ai_composer/validator.py`.

Top-level gates:

- `structural`;
- `semantic`;
- `geometry`;
- `visual`;
- `application`.

Statuses are explicit. Unavailable visual/application gates remain `not_executed`; they never inherit PASS from parser openability.

## C8. `ReviewPacket` — preview/orchestration boundary

**Runtime status:** implemented in `pptx_ai_composer/review_contract.py`; published as `review.json` by `preview.py`.

```json
{
  "contract_version": "1.0",
  "kind": "pptx_chat_review",
  "interaction": "chat_only",
  "deck_id": "deck_demo",
  "revision": "<canonical validated source bundle SHA-256>",
  "fidelity": "structural_preview_not_powerpoint_render",
  "limitations": [],
  "diagnostics": {"text_overflow": []},
  "slides": [
    {
      "slide_id": "problem",
      "number": 1,
      "png_file": "slide-01.png",
      "png_sha256": "...",
      "svg_file": "slide-01.svg",
      "svg_sha256": "..."
    }
  ]
}
```

Invariants:

- the contract is closed recursively and validated before atomic publication;
- `revision` is computed inside the trusted preview boundary and is not caller-overridable: the conversational `render_preview`/CLI path binds the exact validated DeckSpec and compiled preview SceneSpec, while the isolated low-level `render_scene_preview` API binds its validated SceneSpec only;
- neither revision includes conversation text or a mutable file path;
- artifacts are local basenames and are independently hash-bound;
- interaction is chat-only: no direct-edit events, prompts, UI state, coordinates or mutation instructions;
- the host owns chat state, auth, attachment registration and retention;
- preview is evidence only; every conversational revision starts from a newly validated DeckSpec, while the low-level scene API starts from a newly validated SceneSpec.

## Current public runtime surface

JSON-lines CLI actions:

1. `catalog`
2. `preview`
3. `render`
4. `validate`

That is intentionally smaller than the internal primitive set. `preview`/`render` accept only named variants and stable slide IDs for local iteration; primitive geometry is not public. Conversation, research and image-provider calls do not belong in this CLI.

## Change discipline

Any new contract field or archetype requires:

1. a demonstrated user workflow;
2. a failing closed-contract test;
3. exact bounds;
4. deterministic refusal behavior;
5. renderer test proving native/editable output where claimed;
6. validator coverage;
7. README and claim-boundary update.
