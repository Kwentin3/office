# Managed composition library — V1 boundary

## Why

The first dogfood proved that a fixed archetype-only renderer is too narrow, while an LLM-facing coordinate/OOXML API would be too fragile. The bounded hierarchy is:

```text
DeckSpec / SlideSpec
→ managed archetype recipe
→ reusable component
→ bounded primitive
→ PPTX or fast-preview backend
```

## Current managed catalog

`pptx_ai_composer/library.py` is the machine-readable registry. Each component is:

- named;
- bounded;
- classified by kind;
- lifecycle-managed as `experimental`, `stable`, or `deprecated`.

Each archetype declares:

- its component recipe;
- its allowed named variants;
- lifecycle state;
- preview fidelity.

The JSON-lines `catalog` action exposes the registry to the orchestrator and planner.

## Current primitive boundary

The shared runtime `SceneSpec` currently implements only:

- background/surface rectangle;
- rounded card;
- circle/badge/bullet;
- wrapped text;
- image and native-chart nodes;
- source footer.

These are runtime-facing primitives. They are deliberately not accepted as arbitrary LLM coordinates in DeckSpec V1.

## Current archetype recipes

- `cover` → background + title + cover panel + source footer;
- `comparison` → background + title + two-column cards + source footer;
- `chart_with_takeaway` → background + title + chart panel + takeaway card + source footer.
- `process` → background + title + bounded step cards/connectors + source footer;
- `timeline` → background + title + bounded milestone track/cards + source footer;
- `decision_matrix` → background + title + bounded criteria/rating grid + source footer;
- `kpi_grid` → background + title + bounded metric cards + source footer.

The trusted compiler is the only owner of coordinates. Preview and native-PPTX backends consume the same validated SceneSpec and do not implement archetype recipes independently.

## Fast feedback

The `preview` action compiles the same approved DeckSpec into:

- one closed chat-only `review.json` bound to the canonical DeckSpec + compiled SceneSpec SHA-256;
- one SVG per slide;
- one 1280×720 PNG per slide;
- one backend manifest with slide IDs, archetypes, variants, limitations, and bounded text-overflow diagnostics.

The CLI response exposes the ReviewPacket path and absolute PNG `display_artifacts` for a host adapter to register through its native file/media surface instead of adding another Office viewer. The PNG path also exists so an LLM/vision critic can immediately inspect the result after each chat revision. Publication is atomic: a reader never sees a directory containing a mix of old and new slides. Preview artifacts are read-only evidence and never mutate DeckSpec.

## Claim boundary

Preview fidelity is intentionally reported as:

```text
structural_preview_not_powerpoint_render
```

It is useful for judging hierarchy, density, rough wrapping, composition, contrast, and iteration direction. Its text-overflow diagnostic is a deterministic preview-font proxy; it does not prove:

- PowerPoint font metrics or substitution;
- exact text overflow;
- native chart fidelity;
- animations/transitions;
- absence of an application repair dialog.

A later production renderer should produce images from the actual PPTX through PowerPoint or LibreOffice and keep this structural preview as the low-latency first frame.

## Next bounded extension

The next demonstrated workflow may add an image-led composition over the same primitive layer. It should not widen SceneSpec unless the existing `image`, `shape`, and `text` nodes cannot express the required bounded recipe.

Custom compositions should still be recipes from catalog components, not executable callbacks or raw OOXML.
