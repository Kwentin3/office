# Changelog

All notable changes will be documented here.

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
