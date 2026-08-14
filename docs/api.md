# Public API guide

## DOCX — `office_artifact_tool`

```python
from office_artifact_tool import DocxArtifactTool

tool = DocxArtifactTool(workspace)
tool.create(model, output)
tool.inspect(source, view="full", query=None)
tool.plan(snapshot, request)
tool.apply(source, plan, output)
tool.validate(source, before=None, expectations=None)
```

CLI: `office-docx`. Full model/operation examples: `packages/docx/AGENT_SKILL.md`.

## XLSX — `xlsx_artifact_tool`

```python
from xlsx_artifact_tool import XlsxArtifactTool

tool = XlsxArtifactTool(workdir)
tool.create(model, output)
tool.inspect(source, view="summary", sheet=None, range_ref=None, query=None)
tool.plan(snapshot, request)
tool.apply(source, plan, output)
tool.validate(source, before=None)
```

CLI: `office-xlsx` using one JSON request on stdin. Full contract: `packages/xlsx/CONTRACT.md`.

## Preservation-first PPTX — `pptx_artifact_tool`

```python
from pptx_artifact_tool import PptxArtifactTool

tool = PptxArtifactTool(workdir)
tool.inspect(source, view="summary", slide_id=None, query=None)
tool.plan(snapshot, request)
tool.apply(source, plan, output)
tool.create(template, model, output)
tool.validate(source, before=None)
```

CLI: `office-pptx-edit`. Use only native placeholders and explicitly named `slot:*` shapes. Contract: `packages/pptx-editor/CONTRACT.md`.

## Creation-first PPTX — `pptx_ai_composer`

```python
from pptx_ai_composer.contracts import validate_deck_spec
from pptx_ai_composer.renderer import render_deck
from pptx_ai_composer.preview import render_preview
from pptx_ai_composer.validator import validate_presentation

validate_deck_spec(deck_spec)
render_preview(deck_spec, preview_directory, protected_paths=[...])
render_deck(deck_spec, output_pptx, protected_paths=[...])
validate_presentation(output_pptx, deck_spec)
```

CLI: `office-pptx-compose`. Managed archetypes and variants are exposed by its `catalog` action.

## Result handling

DOCX/XLSX/PPTX editor operations return typed dictionaries. Treat `refused` as a normal safe outcome. Never weaken a refusal by constructing a lower-level library call.

The composer CLI uses `status:error`, `code`, and `message` for invalid requests; renderer/preview Python calls raise their documented bounded error types.
