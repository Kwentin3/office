# XLSX domain

Create and preservation-first bounded operations for `.xlsx` workbooks.

## Import

```python
from xlsx_artifact_tool import XlsxArtifactTool
```

## Lifecycle

```text
create → inspect(summary/search/region) → plan → apply → validate
```

Formulas are explicit; a plain value beginning with `=` is refused. The runtime never fabricates calculated values. Existing signed/macro/rich-feature workbooks are refused for mutation when preservation is not proven.

See `CONTRACT.md`, `AGENT_SKILL.md`, and `schemas/`.
