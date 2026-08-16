# Changelog

All notable changes will be documented here.

## Unreleased

- Documented the verified downstream Hermes WebUI five-tool review/export integration, direct-user revision approval boundary, isolated runtime pattern, native `MEDIA:` delivery responsibility, and remaining evidence limits.
- Clarified that `display_artifacts` are host-neutral package results: the Office domains do not emit Hermes `MEDIA:` handles or Open WebUI File/Rich UI events, and this repository does not yet ship either host integration.

## 0.4.0 — 2026-08-15

- Added chat-only creation review loops for DOCX, XLSX and PPTX, with complete validated semantic models as the only source of truth.
- Added closed revision-bound ReviewPackets and host-neutral display artifacts: escaped structural HTML for DOCX/XLSX and hash-bound SVG/PNG for PPTX.
- Added Linux directory-atomic DOCX/XLSX/PPTX review replacement with continuous old-or-new visibility, fail-closed pre-commit behavior and honest best-effort post-commit cleanup.
- Hardened PPTX asset admission with host allowlists, immutable byte snapshots, digest/format/decode/bounds checks and nonblocking special-file refusal.
- Aligned DOCX/XLSX preview admission with final creation, including canonical XLSX structural references and DOCX style/XML-text validation.
- Added request-isolated, descriptor-bound Open WebUI adapter methods returning native `display_artifacts`, plus regression, distribution and clean-wheel coverage.

## 0.3.0 — 2026-08-14

- Added format-local Rich Inspect schema v2 findings with deterministic bounded DOCX story, XLSX worksheet/range and PPTX slide/shape locations.
- Localized only preservation-proven policy decisions; package-global macros, signatures, protection, ActiveX and OLE remain global blockers, and incomplete findings fail closed.
- Added executable preservation matrices covering same/different scope edits and exact non-target ZIP-member equality across DOCX, XLSX and PPTX.
- Added closed host-supplied Application Witness runtime identity without allowing request JSON to choose the executable, work directory or identity.
- Added fixed-argv manual LibreOffice matrix and XLSX cached-value recalculation gates with hermetic fake-executable tests.
- Hardened application evidence against normalized-formula destruction, runtime-identity mutation, cleanup directory swaps and package-root relationship traversal.
- Added a manual-only self-hosted LibreOffice workflow; no schedule or production compatibility claim is enabled without real pinned-runtime evidence.

## 0.2.0 — 2026-08-14

- Added format-specific Rich Inspect inventories and mutation policy decisions for DOCX, XLSX, and PPTX.
- Added strict `{{token}}` template compilation through each format's existing closed plan/apply pipeline.
- Added clone-only `office_application_witness` with fixed argv, private profile, timeout, bounded output validation, cleanup, and typed reports.
- Kept LibreOffice optional and externally supplied; CI validates the operational contract with a hermetic fake executable and does not claim real Office compatibility.
- Added the `office-witness` CLI while preserving document-domain isolation.
- Hardened strict templates against malformed markers, post-render overflow, unsupported package scopes and collateral run/paragraph loss.
- Made the Witness trust boundary explicit: cleanup failure refuses success, while the host-pinned executable remains outside an OS sandbox claim.

## 0.1.0 — 2026-08-13

- Initial public packaging of four isolated bounded Office domains.
- DOCX create/inspect/plan/apply/validate tool.
- XLSX create/inspect/plan/apply/validate tool.
- Preservation-first PPTX template editor.
- Creation-first managed PPTX composer and structural preview.
- GitHub Actions, integration documentation, schemas, agent skills, and examples.
- Hardened preservation-first PPTX apply-side preflight for expected text, text budgets, and mixed reorder/content plans.
- Bound preservation-first PPTX apply to a private immutable source snapshot, closing source-path replacement TOCTOU.
- Enforced documented PPTX cardinality budgets for slides, slots, and table cells at both inspection and forged planning boundaries.
- Added symlink-safe, exclusive, descriptor-relative per-request staging tailored to the `corp-openweb-ui` Linux sidecar layout, including concurrent directory-swap detection and cleanup.
- Added a fail-closed wheel/sdist inventory gate with regression tests for every required runtime resource, all four CLI entry points, Apache notices, and bundled font licensing.
