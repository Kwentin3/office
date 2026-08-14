# Security policy

## Supported version

The latest `main` revision is the only supported development version until the first stable release.

## Reporting

Please report suspected vulnerabilities privately through GitHub's security-advisory feature. Do not attach sensitive Office documents to public issues.

## Security boundary

These tools enforce bounded contracts, package admission, private candidates, semantic postconditions, and atomic publication. They are **not** a complete multi-tenant sandbox. A host application must still enforce:

- per-user/request workspace isolation;
- upload size and ownership policy;
- authentication and authorization;
- retention and deletion policy;
- malware/content scanning when required;
- Office/LibreOffice application-level compatibility checks.

Macro-enabled, signed, encrypted, and unsupported rich-feature documents are refused or treated as read-only according to each domain contract.
