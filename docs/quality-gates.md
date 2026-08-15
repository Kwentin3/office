# Quality gates

## v0.4.0 bounded evidence

| Domain | Local suite | Independent exact-tree gate | Application evidence |
|---|---:|---|---|
| DOCX | 56/56 | required before tag | preservation matrix + fake-soffice contract + chat review |
| XLSX | 75/75 | required before tag | preservation matrix + cached-value gate + chat review |
| PPTX editor | 44/44 | required before tag | preservation matrix + fake-soffice contract |
| PPTX composer | 92/92 | required before tag | hash-bound SVG/PNG structural preview only |
| Application Witness | 28/28 | required before tag | hermetic fake executable |
| Repository integration | 40/40 | required before tag | descriptor-bound WebUI adapter + manual workflow contract |

`scripts/test_all.py` passes **335/335** unittest cases. The authoritative cross-package pytest run passes **352/352 tests plus 314 subtests**. Repository CI runs both suites on Python 3.11 and 3.12, then verifies the built archives and a clean wheel installation.

Distribution evidence for the v0.4.0 candidate:

- wheel and sdist built into an external temporary directory from the candidate tree;
- fail-closed verifier passed with 70 wheel members and 198 sdist members;
- all five console scripts, three packaged inventory schemas, domain imports and clone-only Witness remain required by the verifier;
- changed/new v0.4.0 modules pass targeted Ruff and repository-wide compileall;
- whole-repository Ruff remains historical style debt and is not a CI/release gate for this bounded feature release.

## What a local PASS means

- closed contracts and typed refusals;
- source/output collision protection;
- source fingerprint and plan binding;
- private candidate and validation before publication;
- bounded package admission;
- tested semantic postconditions;
- source remains unchanged in tested flows;
- Rich Inspect schema v2 provides bounded, deterministic, format-local locations for listed features;
- package-global macros, signatures, protection, ActiveX and OLE remain global blockers;
- preservation matrices prove exact non-target archive-member equality for the tested same/different-scope flows;
- Application Witness observes a private clone, deletes its private workspace before success and never publishes its output;
- runtime identity is closed and host-supplied, not request-controlled;
- the Witness controls paths and process-group lifetime but explicitly does not claim OS sandboxing of the host-pinned executable;
- the XLSX semantic gate verifies only declared cached cells with explicit type/tolerance, exact formula preservation and unsupported-formula classification;
- application-gate cleanup and reported runtime identity are revalidated before a success result.

## Unexecuted production gates

- execution of `.github/workflows/libreoffice-compat.yml` on an actual pinned self-hosted LibreOffice runtime;
- a stable real-LibreOffice evidence series; no nightly schedule is enabled;
- real Microsoft Office open-save-reopen;
- repair-dialog detection;
- pixel/render fidelity and font substitution;
- Word pagination;
- hostile multi-tenant deployment controls.

Therefore the honest status remains:

```text
bounded release rule: exact-tree APPROVE and GitHub CI are mandatory before tag
production compatibility: HOLD until host-specific application gates
```

A successful fake-soffice or fake-recalculation suite proves subprocess lifecycle, private-path handling, declared cached-value comparison and reporting contracts only. It is not OS-sandbox or real application-compatibility evidence and never implies Microsoft Office equivalence.
