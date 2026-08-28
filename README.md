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
python -m pip install "kwentin-office @ https://github.com/Kwentin3/office/archive/refs/tags/v0.5.0.zip"
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

Each call atomically publishes a closed `review.json` plus escaped, script-free HTML review artifacts and returns absolute `display_artifacts` for a host adapter to register and display. DOCX resolves one immutable `professional-a4/v2` presentation contract into both its styled page-layout proxy and final DOCX, including page geometry, managed typography, spacing, lists, and table geometry. Its revision binds the validated semantic model plus that presentation ID. This is still not Word pagination: line wrapping, font substitution, and page breaks remain browser approximations. XLSX remains a bounded structural grid that neither executes formulas nor claims Excel layout fidelity. Replacing an existing review requires Linux `renameat2(RENAME_EXCHANGE)` support and fails closed without changing the old review when that primitive is unavailable. Chat revises the complete semantic model; preview never becomes editable state.

### Creation-first PPTX

```bash
office-pptx-compose <<'JSON'
{"action":"preview","spec":"examples/pptx-composer/managed-library.deck.json","output":"deck-preview"}
JSON
```

`preview` atomically publishes a closed chat-only `review.json` and per-slide SVG/PNG files. Replacing an existing preview uses Linux `renameat2(RENAME_EXCHANGE)`, matching the DOCX/XLSX continuous old-or-new visibility boundary. Its JSON response exposes host-neutral `display_artifacts`; a host adapter must register and deliver those PNGs through the host's native file/media surface. Chat revises the DeckSpec, then calls `preview` again. There is no direct preview editing or second document state. Image nodes appear as labeled placeholders in this structural preview; their admitted pixels are embedded only by the final native PPTX renderer.

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

## Integration into Open WebUI

Recommended boundary, verified against the official Open WebUI extension documentation:

```text
Open WebUI chat
→ external Office OpenAPI sidecar
→ authenticated per-request workspace
→ exactly one domain package call
→ typed result/refusal
→ optional thin native Tool/Action registers File objects or persisted Rich UI
```

Use OpenAPI as the service boundary, not legacy Pipelines. External tools can use Open WebUI's authenticated endpoint for supported one-way UI events. Keep any native Workspace Tool/Action limited to host-local file handling, native Rich UI conversion, bidirectional confirmation/input, and authenticated sidecar calls; Workspace Tools execute inside the Open WebUI process and must not contain Office mutation logic. `MEDIA:` is Hermes-specific and is not an Open WebUI transport contract. See the integration guide for the checked upstream sources, the documented event/MCP wording tension, reserved-argument stability, and revision-bound approval persistence.

The repository currently provides the Office packages and path-safe `OfficeService` example, not a ready-to-register FastAPI/OpenAPI server or native Open WebUI Tool/Action.

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
