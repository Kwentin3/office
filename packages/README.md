# Domain packages

Each directory is an independently bounded domain included in the single `kwentin-office` distribution.

| Directory | Import package | CLI | Purpose |
|---|---|---|---|
| `docx` | `office_artifact_tool` | `office-docx` | DOCX create and preservation-first edits |
| `xlsx` | `xlsx_artifact_tool` | `office-xlsx` | XLSX create, cells, formulas, rows, styles |
| `pptx-editor` | `pptx_artifact_tool` | `office-pptx-edit` | Template/slot-based preservation-first PPTX editing |
| `pptx-composer` | `pptx_ai_composer` | `office-pptx-compose` | Creation-first semantic deck composition |

There is intentionally no shared `OfficeArtifactBase`, Office AST, OOXML runtime, or cross-format SceneSpec.
