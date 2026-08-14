# Application Witness package

Import: `office_application_witness`
CLI: `office-witness`

This package observes a private clone with a host-pinned trusted LibreOffice executable. It has no Office AST, cannot mutate through any domain backend, retains no private-workspace output after success, and never claims Microsoft Office equivalence. It controls subprocess lifetime and paths but is not an OS sandbox.

See [`../../docs/application-witness.md`](../../docs/application-witness.md).
