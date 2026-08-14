# Contributing

## Principles

This repository shares transaction **policy**, not a cross-format domain model.

- Keep DOCX, XLSX, preservation-first PPTX, and creation-first PPTX isolated.
- Do not add raw OOXML, backend objects, arbitrary Python, callbacks, or unrestricted coordinates to public contracts.
- Add a failing regression test before fixing a defect.
- Build a private candidate, validate it independently, then publish atomically.
- Never overwrite a source artifact.

## Development setup

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install -e '.[dev]'
python scripts/test_all.py
```

## Pull requests

1. Keep changes within one domain unless the policy is genuinely identical.
2. Document contract changes.
3. Run `python scripts/test_all.py` and `python -m build`.
4. State which application-level gates were not executed.

## Commit style

Use Conventional Commits, for example:

```text
fix(xlsx): bind reordered row payloads to source snapshot
```
