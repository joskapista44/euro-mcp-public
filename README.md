# Euro-MCP

Euro-MCP is an MCP server for reliable, agent-driven Office document work through the Euro-Office / ONLYOFFICE Document Server stack.

It gives an AI agent a verified path to read, create, edit, format, and persist Office documents without silently falling back to a local Office engine or direct OOXML rewriting.

## v1.0.0: agent-grade XLSX editing

The first release introduces a persistent live XLSX execution path designed for complex, multi-step spreadsheet tasks.

A batch is executed in one editor session and follows an evidence-based lifecycle:

1. inspect the live workbook;
2. resolve and validate targets;
3. apply the requested mutations;
4. read the affected state back from the live editor;
5. verify the result against the request;
6. cross a wrapper-owned persistence barrier;
7. support a persisted reopen retry with zero writes when the task is already satisfied.

The public result distinguishes live observation from live verification through the `LIVE_READ` and `LIVE_VERIFY` authorities. Failures are reported with operation-bound diagnostics instead of being treated as successful writes.

### XLSX power-user surface

The v1.0.0 XLSX path covers the main spreadsheet operation families:

- worksheets, ranges, values, formulas, clearing, merging, and defined names;
- formatting, borders, number formats, conditional formatting, AutoFit, row height, and column width;
- sorting, filtering, data validation, and freeze panes;
- row and column insertion/deletion plus range and sheet copy/move operations;
- charts and pivot-table workflows supported by the deployed live editor surface;
- page layout, paper size, orientation, margins, scaling, print area, print titles, page breaks, first-page numbering, centering, and headers/footers;
- mixed cross-family tasks executed and verified in one persistent live session.

A request such as “fit the final worksheet onto one A4 page” is supported at the spreadsheet page-setup level: the agent can set A4 paper dimensions, orientation, print area, margins, and fit-to-width/fit-to-height scaling, then verify and persist those settings.

### Acceptance evidence

The release is backed by static contract tests and TRUE LIVE ONLYOFFICE acceptance runs, including:

- a cross-family 19-operation task completed with `LIVE_VERIFY`, 19 checks, four readbacks, and a successful save barrier;
- a persisted reopen of the same task completed as a zero-write no-op with all 19 checks satisfied;
- integrated page/print acceptance, including A4 one-page setup and persisted zero-write retry;
- operation-bound batch receipts and failure diagnostics.

See [XLSX-AGENT-GRADE-MIGRATION-STATUS.md](XLSX-AGENT-GRADE-MIGRATION-STATUS.md), [EXCEL-POWER-USER-CAPABILITY-GAPS.md](EXCEL-POWER-USER-CAPABILITY-GAPS.md), and [docs/xlsx-page-print-persistent-live-closure.md](docs/xlsx-page-print-persistent-live-closure.md) for the detailed evidence and current boundaries.

### Explicit boundaries

“Power user” means the documented and accepted operation surface, not arbitrary parity with every Excel desktop feature. Deferred or environment-dependent areas remain fail-closed and are recorded in the canonical capability-gap register. Examples include full ListObject/table identity, chart copying, complete sheet-password protection semantics, full protected-range lifecycle identity, broad pivot equivalence, and rendered PDF page-count verification.

## Architecture boundary

Document operations belong on Euro-Office. This repository does **not** contain or use a local Office engine, headless LibreOffice/`soffice`, or direct on-disk OOXML editing as an alternative execution path.

The supported paths are:

- `euro-mcp.cjs` — MCP entry point and tool surface;
- `runner.cjs` + `box-helper.py` — remote DocBuilder job transport and execution on the Euro-Office box;
- `coedit.cjs` — operations against a live co-editing session;
- `lib.cjs`, `lib-operations-*.cjs`, and `operations/*.cjs` — ONLYOFFICE/DocBuilder operation generation and translation;
- `xlsx-persistent-*.cjs`, `xlsx-agent-*.cjs`, and `xlsx-*-observer.cjs` — persistent live XLSX planning, execution, observation, and verification;
- `package-consistency.cjs` — validation of Document Server-produced OOXML packages;
- `capabilities-registry-build.mjs` and `capability-status-model.mjs` — capability evidence/status tooling.

If a capability cannot be implemented and verified through the Euro-Office / Document Server path, it is reported as unavailable until that path exists. A local document-processing fallback is intentionally out of scope.

## Requirements

- Node.js 22 or newer;
- Python 3 for `box-helper.py` and its tests; no third-party Python packages are required;
- access to the Euro-Office DocBuilder box for the remote execution path;
- a Nextcloud + ONLYOFFICE deployment for live co-editing operations;
- Playwright for co-editing where required by the deployment. It may be supplied through normal Node resolution or `EURO_PLAYWRIGHT_PATH`.

## Install

```sh
npm ci
cp .env.example .env
```

Fill only the settings needed by your deployment. Secrets must stay out of the repository.

## Configuration

The main configuration groups are:

- `EURO_EXEC`, `EURO_SSH_*`, `EURO_TIMEOUT_MS`, `EURO_BOX_IP` — transport to the DocBuilder box;
- `EURO_AGENT_ID` — caller identity used for audit/trace context;
- `EURO_COEDIT_*` and `<AGENT>_NEXTCLOUD_APP_PASSWORD` — live co-editing access;
- `EO_DS_URL`, `EO_ENV_FILE` — box-side DocBuilder/Document Server configuration;
- `EURO_TRACE_LOG_DIR` — optional trace/capability diagnostics.

See `.env.example` for safe placeholders and comments.

## Testing

```sh
npm test
```

The standard test command runs the repository-contained JavaScript/MJS suites plus the `box-helper.py` unit tests. CI is designed not to require production credentials or a live customer document.

Representative core gates:

```sh
node test-lib.cjs
node test-tools.cjs
python3 -m unittest discover -p 'test_box_helper_*.py'
```

Tests named `*_live_acceptance.cjs` are environment-bound acceptance gates and require the configured Euro-Office / ONLYOFFICE runtime.

## CI architecture guard

CI enforces the product boundary. It rejects reintroduction of the retired offline toolkit, local LibreOffice/`soffice` execution settings, and known local document-engine modules. This makes the Document Server-only rule executable rather than documentation-only.

## Project layout

- `euro-mcp.cjs`, `runner.cjs`, `coedit.cjs`, `box-helper.py` — MCP, transport, and Euro-Office integration;
- `lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs`, `euro-magok.cjs` — document operations implemented for ONLYOFFICE/DocBuilder;
- `xlsx-agent-*.cjs`, `xlsx-persistent-*.cjs`, `xlsx-*-observer.cjs` — agent-grade XLSX execution and verification;
- `office-trace.cjs`, `capabilities-registry-build.mjs`, `capability-status-model.mjs` — observability and capability evidence;
- `package-consistency.cjs` — output package integrity validation;
- `EXCEL-POWER-USER-CAPABILITY-GAPS.md` — canonical XLSX deferred/unsupported capability register;
- `test*.cjs`, `test*.mjs`, `test_box_helper_*.py` — automated and live acceptance tests.

## Security

See [SECURITY.md](SECURITY.md) for the vulnerability-reporting process and the security-sensitive boundaries in the Euro-Office integration.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).
