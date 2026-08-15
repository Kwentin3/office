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
python -m pip install "kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.4.0.zip"
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
| DOCX | `office_artifact_tool` | `office-docx` | create + chat review + preservation-first edit |
| XLSX | `xlsx_artifact_tool` | `office-xlsx` | create + chat review + cell/range/formula edit |
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

### Chat-only DOCX/XLSX review

Creation-first Word and Excel work can be reviewed before the Office export:

```python
from office_artifact_tool import render_docx_preview
from xlsx_artifact_tool import render_xlsx_preview

word_review = render_docx_preview(docx_model, "docx-review")
excel_review = render_xlsx_preview(xlsx_model, "xlsx-review")
```

Each call atomically publishes a closed `review.json` plus escaped, script-free HTML structural previews and returns absolute `display_artifacts` for native Hermes WebUI `MEDIA:` rendering. Replacing an existing review requires Linux `renameat2(RENAME_EXCHANGE)` support and fails closed without changing the old review when that primitive is unavailable. Chat revises the complete semantic model; preview never becomes editable state. DOCX preview is not Word pagination, and XLSX preview neither executes formulas nor claims Excel layout fidelity. After approval, pass the exact revised model to the existing `create(...)` method to produce the Office file.

### Creation-first PPTX

```bash
office-pptx-compose <<'JSON'
{"action":"preview","spec":"examples/pptx-composer/managed-library.deck.json","output":"deck-preview"}
JSON
```

`preview` atomically publishes a closed chat-only `review.json` and per-slide SVG/PNG files. Replacing an existing preview uses Linux `renameat2(RENAME_EXCHANGE)`, matching the DOCX/XLSX continuous old-or-new visibility boundary. Its JSON response exposes `display_artifacts`; Hermes WebUI can show those PNG paths through native `MEDIA:` rendering. Chat revises the DeckSpec, then calls `preview` again. There is no direct preview editing or second document state. Image nodes appear as labeled placeholders in this structural preview; their admitted pixels are embedded only by the final native PPTX renderer.

After approval:

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

- execution of the manual self-hosted LibreOffice application gate on a host-pinned runtime (the gate and its hermetic contract tests ship, but LibreOffice is not installed by this package);
- renderer-based visual regression and font-substitution checks;
- real Excel/LibreOffice formula recalculation evidence from the declared cached-value gate;
- multi-tenant host sandboxing and retention policy.

See [quality gates](docs/quality-gates.md).

## License

Apache License 2.0. The bundled Noto Sans preview font is covered by the SIL Open Font License; see `NOTICE` and its adjacent `OFL.txt`.
