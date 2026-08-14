# Quality gates

## v0.2.0 bounded evidence

| Domain | Local suite | Independent exact-tree gate | Application evidence |
|---|---:|---|---|
| DOCX | 49/49 | required before tag | fake-soffice contract only |
| XLSX | 51/51 | required before tag | fake-soffice contract only |
| PPTX editor | 36/36 | required before tag | fake-soffice contract only |
| PPTX composer | 66/66 | unchanged creation-first domain | structural preview only |
| Application Witness | 24/24 | required before tag | hermetic fake executable |
| Repository integration | 15/15 | required before tag | N/A |

Current local total: **241/241 tests passed**. The repository CI reruns the packaged source rather than trusting these counts.

Distribution evidence:

- wheel and sdist built into an external temporary directory from the candidate tree;
- fail-closed verifier passed with 63 wheel members and 175 sdist members;
- clean external venv installed `kwentin-office==0.2.0` from the wheel;
- all five console scripts, three packaged inventory schemas, domain imports and clone-only Witness smoke passed;
- new v0.2.0 modules pass Ruff and compileall;
- whole-repository Ruff remains historical style debt and is not a CI/release gate for this bounded feature release.

## What a local PASS means

- closed contracts and typed refusals;
- source/output collision protection;
- source fingerprint and plan binding;
- private candidate and validation before publication;
- bounded package admission;
- tested semantic postconditions;
- source remains unchanged in tested flows;
- Rich Inspect is a bounded inventory of listed features, not a complete OOXML parser;
- Application Witness observes a private clone, deletes its private workspace before success and never publishes its output;
- the Witness controls paths and process-group lifetime but explicitly does not claim OS sandboxing of the host-pinned executable.

## Unexecuted production gates

- real Microsoft Office or LibreOffice open-save-reopen across supported versions;
- repair-dialog detection;
- pixel/render fidelity and font substitution;
- Word pagination;
- verified Excel formula recalculation and compatibility;
- hostile multi-tenant deployment controls.

Therefore the honest status remains:

```text
bounded release rule: exact-tree APPROVE and GitHub CI are mandatory before tag
production compatibility: HOLD until host-specific application gates
```

A successful fake-soffice suite proves subprocess lifecycle, private-path handling and reporting contracts only. It is not OS-sandbox or application-compatibility evidence and never implies Microsoft Office equivalence.
