# Contributing

## Before you start

Euro-MCP has one architectural boundary: document operations run through Euro-Office / OnlyOffice Document Server. Do not add a local Office engine, LibreOffice/`soffice` fallback, or direct on-disk OOXML editing path as an alternative execution engine.

If the Document Server path cannot support a capability yet, keep that capability unavailable or extend the Euro-Office integration rather than bypassing it locally.

## Workflow

1. Keep changes scoped to one fix or capability per pull request.
2. Add or update tests for every behavior change.
3. Run `npm test` before requesting review.
4. Keep fake-backed tests credential-free; production secrets and customer documents must never be required by CI.
5. Call out changes to transport, co-editing authorization, credential handling, or output-package validation explicitly in the PR description.

## Code style

- Node/CJS is the primary runtime; there is no bundling/build step for the MCP server.
- Keep dependencies small and justify new runtime dependencies.
- Python is retained for the remote `box-helper.py` path and should remain standard-library-only unless there is a deliberate design change.
- Match the naming and error-message language of the file being edited.

## Architecture-sensitive changes

Changes touching these areas deserve extra review:

- `coedit.cjs` caller identity and `EURO_COEDIT_AGENTS` authorization;
- `runner.cjs` SSH key and remote execution handling;
- `box-helper.py` Document Server/DocBuilder execution contract;
- `lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs` OnlyOffice operation generation;
- `package-consistency.cjs` output integrity checks;
- the CI architecture guard that prevents local document-engine paths from returning.

See `SECURITY.md` for private vulnerability reporting.
