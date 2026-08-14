# DOCX domain

Create and preservation-first bounded edits for `.docx` files.

## Import

```python
from office_artifact_tool import DocxArtifactTool
```

## Lifecycle

```text
create(model, output)
inspect(source, view, query)
plan(snapshot, request)
apply(source, plan, output)
validate(source, before, expectations)
```

The source is never overwritten. Plans and transaction IDs are bound to the exact source snapshot. Candidate package and semantic checks run before atomic publication.

See `AGENT_SKILL.md` for model-facing envelopes and `schemas/` for JSON contracts.
