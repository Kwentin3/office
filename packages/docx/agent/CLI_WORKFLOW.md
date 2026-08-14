# CLI quickstart

All payloads are JSON files. The model does not write Python.

```bash
PYTHONPATH=. python -m office_artifact_tool --workspace ./work create \
  --model model.json --output created.docx

PYTHONPATH=. python -m office_artifact_tool --workspace ./work inspect \
  --source source.docx > snapshot.json

PYTHONPATH=. python -m office_artifact_tool --workspace ./work plan \
  --snapshot snapshot.json --request request.json > planned.json

# Extract the `plan` field into plan.json.
PYTHONPATH=. python -m office_artifact_tool --workspace ./work apply \
  --source source.docx --plan plan.json --output changed.docx

PYTHONPATH=. python -m office_artifact_tool --workspace ./work validate \
  --source changed.docx
```

Expected refusals are JSON with `status`, `reason`, and `details`. CLI exit code is zero for a well-formed safe refusal and nonzero for an invalid result or process failure.
