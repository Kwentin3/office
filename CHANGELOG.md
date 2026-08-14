# Changelog

All notable changes will be documented here.

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
