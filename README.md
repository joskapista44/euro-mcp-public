# Euro-MCP

Euro-MCP is an MCP server for reading, creating and editing Office documents through the Euro-Office / OnlyOffice Document Server stack.

## Architecture boundary

Document operations belong on Euro-Office. This repository does **not** contain or use a local Office engine, headless LibreOffice/`soffice`, or direct on-disk OOXML editing as an alternative execution path.

The supported paths are:

- `euro-mcp.cjs` — MCP entry point and tool surface.
- `runner.cjs` + `box-helper.py` — remote DocBuilder job transport and execution on the Euro-Office box.
- `coedit.cjs` — operations against a live co-editing session.
- `lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs` — OnlyOffice/DocBuilder operation generation and translation.
- `package-consistency.cjs` — validates Document Server-produced OOXML packages before they are returned.
- `capabilities-registry-build.mjs` and `capability-status-model.mjs` — capability evidence/status tooling.

If a capability cannot be implemented through the Euro-Office / Document Server path, it must be reported as unavailable until that path exists. A local document-processing fallback is intentionally out of scope.

## Requirements

- Node.js 22 or newer.
- Python 3 for `box-helper.py` and its tests; no third-party Python packages are required.
- Access to the Euro-Office DocBuilder box for the remote execution path.
- A Nextcloud + OnlyOffice deployment for co-editing operations.
- Playwright for co-editing where required by the deployment. It may be supplied through normal Node resolution or `EURO_PLAYWRIGHT_PATH`.

## Install

```sh
npm ci
cp .env.example .env
```

Fill only the settings needed by your deployment. Secrets must stay out of the repository.

## Configuration

The main configuration groups are:

- `EURO_EXEC`, `EURO_SSH_*`, `EURO_TIMEOUT_MS`, `EURO_BOX_IP` — transport to the DocBuilder box.
- `EURO_AGENT_ID` — caller identity used for audit/trace context.
- `EURO_COEDIT_*` and `<AGENT>_NEXTCLOUD_APP_PASSWORD` — live co-editing access.
- `EO_DS_URL`, `EO_ENV_FILE` — box-side DocBuilder/Document Server configuration.
- `EURO_TRACE_LOG_DIR` — optional trace/capability diagnostics.

See `.env.example` for safe placeholders and comments.

## Running the tests

```sh
npm test
```

The test command runs the repository-contained JavaScript/MJS suites plus the `box-helper.py` unit tests. The CI suite is designed not to require production credentials or a live customer document.

For individual core gates:

```sh
node test-lib.cjs
node test-tools.cjs
python3 -m unittest discover -p 'test_box_helper_*.py'
```

## CI architecture guard

CI also enforces the product boundary. It rejects reintroduction of the retired offline toolkit, local LibreOffice/`soffice` execution settings, and known local document-engine modules. This makes the Document Server-only rule executable rather than documentation-only.

## Project layout

- `euro-mcp.cjs`, `runner.cjs`, `coedit.cjs`, `box-helper.py` — MCP, transport and Euro-Office integration.
- `lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs`, `euro-magok.cjs` — document operations implemented for OnlyOffice/DocBuilder.
- `office-trace.cjs`, `capabilities-registry-build.mjs`, `capability-status-model.mjs` — observability and capability evidence.
- `package-consistency.cjs` — output package integrity validation.
- `test*.cjs`, `test*.mjs`, `test_box_helper_*.py` — automated tests.

## Security

See `SECURITY.md` for the vulnerability-reporting process and the security-sensitive boundaries in the Euro-Office integration.

## Contributing

See `CONTRIBUTING.md`.
