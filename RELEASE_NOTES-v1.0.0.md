# Euro-MCP v1.0.0 — Agent-grade XLSX editing

Euro-MCP v1.0.0 is the first public release.

This release introduces a persistent, TRUE LIVE ONLYOFFICE execution path for complex agent-driven spreadsheet work. Multi-step tasks are planned, applied, read back, verified, and persisted in one editor session, with operation-bound receipts and fail-closed diagnostics.

## Highlights

- Persistent one-editor-session XLSX batches.
- Live target discovery and precondition checks before mutation.
- Post-mutation readback with `LIVE_READ` / `LIVE_VERIFY` authority.
- Wrapper-owned persistence barrier.
- Persisted reopen retry that performs zero writes when the requested state is already satisfied.
- Mixed cross-family tasks across data, structure, formatting, layout, and print settings.

## XLSX power-user capabilities

- Worksheet and range creation/editing.
- Values, formulas, clearing, merging, and defined names.
- Number formats, borders, conditional formatting, AutoFit, row height, and column width.
- Sorting, filtering, data validation, and freeze panes.
- Row/column insertion and deletion.
- Range and sheet copy/move workflows.
- Charts and supported pivot-table workflows.
- Page size, orientation, margins, scaling, print area, print titles, page breaks, first-page numbering, centering, and headers/footers.
- A4 one-page page setup using A4 paper dimensions plus fit-to-width and fit-to-height scaling.

## Acceptance evidence

The release state passed repository static gates and TRUE LIVE ONLYOFFICE acceptance.

The final cross-family acceptance completed:

- `LIVE_VERIFY`;
- 19 operation checks;
- four live readbacks;
- 19 writes in the first application;
- a successful persistence barrier;
- a persisted reopen retry with zero writes and all 19 checks satisfied.

Integrated page/print acceptance also covers A4 one-page setup and persisted zero-write retry.

Detailed evidence:

- [XLSX-AGENT-GRADE-MIGRATION-STATUS.md](XLSX-AGENT-GRADE-MIGRATION-STATUS.md)
- [EXCEL-POWER-USER-CAPABILITY-GAPS.md](EXCEL-POWER-USER-CAPABILITY-GAPS.md)
- [docs/xlsx-page-print-persistent-live-closure.md](docs/xlsx-page-print-persistent-live-closure.md)

## Architecture and safety

Office work remains on the Euro-Office / ONLYOFFICE Document Server path. This project does not silently use a local LibreOffice/`soffice` engine or direct on-disk OOXML editing as a fallback.

Unsupported or unverifiable operations fail closed instead of being reported as successful.

## Known boundaries

The v1.0.0 “power-user” status applies to the documented, accepted operation surface. It does not claim arbitrary parity with every Excel desktop feature.

Current deferred or runtime-dependent areas include:

- full ListObject/table identity;
- chart copying;
- complete sheet-password protection semantics;
- full protected-range lifecycle identity;
- broad pivot equivalence;
- verification of the final rendered PDF page count.

The canonical current list is maintained in [EXCEL-POWER-USER-CAPABILITY-GAPS.md](EXCEL-POWER-USER-CAPABILITY-GAPS.md).
