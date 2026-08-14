# Kwentin Office

Contract-first Python tools for creating, inspecting, editing, validating, and independently observing Office artifacts from applications and LLM runtimes.

## Why this exists

A typical integration asks an LLM to generate one-off `python-docx`, `openpyxl`, or `python-pptx` code. That makes target selection ambiguous, preservation accidental, and validation optional. This repository replaces that pattern with a closed lifecycle:

```text
closed JSON/Python contract
→ deterministic plan or compiler
→ private candidate
→ package + semantic validation
→ atomic publication
```

## Install

From the versioned GitHub archive (works in slim containers without a `git` binary):

```bash
python -m pip install "kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.2.0.zip"
```

For development:

```bash
git clone https://github.com/Kwentin3/office.git
cd office
python -m venv .venv
. .venv/bin/activate
python -m pip install -e '.[dev]'
python scripts/test_all.py
```

Python 3.11+ is required.

## Package map

| Domain | Python import | CLI | Mode |
|---|---|---|---|
| DOCX | `office_artifact_tool` | `office-docx` | create + preservation-first edit |
| XLSX | `xlsx_artifact_tool` | `office-xlsx` | create + cell/range/formula edit |
| PPTX editor | `pptx_artifact_tool` | `office-pptx-edit` | preservation-first template/slot edit |
| PPTX composer | `pptx_ai_composer` | `office-pptx-compose` | creation-first semantic composition |
| Application witness | `office_application_witness` | `office-witness` | clone-only LibreOffice observation |

One distribution makes installation simple. The four document runtimes remain isolated: no common Office AST, no shared OOXML runtime, and no cross-format SceneSpec. The witness is an operational subprocess boundary, not a document runtime.

## Quick start

### DOCX

```python
from pathlib import Path
from office_artifact_tool import DocxArtifactTool

tool = DocxArtifactTool(Path(".office-work/docx"))
result = tool.create(
    {"blocks": [
        {"type": "heading", "level": 1, "text": "Report"},
        {"type": "paragraph", "text": "Validated output", "style": "Normal"},
    ]},
    Path("report.docx"),
)
assert result["status"] == "ok"
```

### XLSX

```python
from pathlib import Path
from xlsx_artifact_tool import XlsxArtifactTool

tool = XlsxArtifactTool(Path(".office-work/xlsx"))
result = tool.create(
    {"sheets": [{"name": "Data", "cells": {
        "A1": {"value": "Item", "style": "header"},
        "B1": {"value": "Amount", "style": "header"},
        "A2": {"value": "A"},
        "B2": {"value": 10, "style": "currency"},
    }}]},
    Path("report.xlsx"),
)
assert result["status"] == "ok"
```

### Creation-first PPTX

```bash
office-pptx-compose <<'JSON'
{"action":"render","spec":"examples/pptx-composer/managed-library.deck.json","output":"deck.pptx"}
JSON
```

### Preservation-first edits

Existing DOCX/XLSX/PPTX files use the same host-controlled pattern:

```text
inspect → choose exact tx_* target → plan → apply to a different path → validate
```

Never invent or persist transaction IDs. They are bound to the exact source fingerprint.

## Integration into `openweb.ui`

Recommended boundary:

```text
Open WebUI route/tool
→ authenticated per-request workspace
→ Python package call
→ typed result/refusal
→ register validated output as an attachment
```

See:

- [Open WebUI integration guide](docs/openwebui-integration.md)
- [host adapter example](examples/openwebui_backend/office_service.py)
- [public APIs](docs/api.md)
- [architecture](docs/architecture.md)
- [Rich Inspect and strict templates](docs/rich-inspect-and-templates.md)
- [Application Witness](docs/application-witness.md)

The core package has no dependency on Open WebUI, FastAPI, a model provider, or a storage vendor.

## Agent contracts

Each domain includes its compact agent instructions and, where applicable, JSON schemas:

- `packages/docx/AGENT_SKILL.md`
- `packages/xlsx/AGENT_SKILL.md`
- `packages/pptx-editor/AGENT_SKILL.md`
- `packages/*/schemas/`

The model should emit only closed requests. It should never receive raw OOXML or generate arbitrary mutation code.

## Validation status

The bounded implementations pass their local regression suites and independent exact-reproducer reviews. User-opened dogfood artifacts were also visually inspected successfully. This is not a claim of complete Office compatibility.

Still required for a strict production sign-off:

- real LibreOffice and Microsoft Word/Excel/PowerPoint application gates (the packaged witness contract is hermetically tested, but LibreOffice is not installed by this package);
- renderer-based visual regression and font-substitution checks;
- Excel/LibreOffice formula recalculation;
- multi-tenant host sandboxing and retention policy.

See [quality gates](docs/quality-gates.md).

## License

Apache License 2.0. The bundled Noto Sans preview font is covered by the SIL Open Font License; see `NOTICE` and its adjacent `OFL.txt`.
