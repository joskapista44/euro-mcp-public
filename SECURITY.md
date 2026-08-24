# Security

## Reporting a vulnerability

Please do not open a public issue for a security vulnerability. Report it privately to the project owner through the repository hosting platform, including the affected files, a minimal reproduction when possible, expected versus actual behavior, and your assessment of impact.

## Security-sensitive boundaries

- **Co-editing identity/authorization** (`coedit.cjs`, `EURO_COEDIT_AGENTS`): controls who may write into a live shared document session.
- **Remote execution and credentials** (`runner.cjs`): handles the SSH identity used to reach the Euro-Office DocBuilder box. Keys and secrets must not be logged or persisted beyond the intended call lifecycle.
- **Box-side execution** (`box-helper.py`): forms the boundary between MCP requests and Document Server/DocBuilder execution on the Euro-Office box.
- **Document operation generation** (`lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs`): malformed or over-broad operations can modify more document state than requested.
- **OOXML package consistency** (`package-consistency.cjs`): validates that Document Server-produced packages are structurally consistent before they are returned.
- **Architecture boundary**: a local document editor, LibreOffice/`soffice` path, or direct file-rewrite fallback is intentionally out of scope. Reintroducing such a path bypasses the expected Euro-Office control boundary and should be treated as an architectural/security regression.

## Dependencies

Bugs in upstream dependencies such as `@modelcontextprotocol/sdk`, `zod`, Playwright, Nextcloud, or OnlyOffice Document Server should normally be reported upstream. If it is unclear whether the issue is in Euro-MCP or a dependency, report it here first.
