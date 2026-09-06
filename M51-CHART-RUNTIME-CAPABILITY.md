# M5.1 chart runtime capability — EuroOffice 9.3.4

Measured live on 2026-09-06 through the production co-edit transport:

`MCP -> Playwright -> Elliot authenticated Nextcloud session -> ONLYOFFICE spreadsheeteditor -> editor.callCommand() -> Api.*`

## Live co-edit capabilities

PASS / machine-read-back verified:

- `ApiWorksheet.AddChart`
- `ApiWorksheet.GetAllCharts`
- `ApiWorksheet.GetAllDrawings`
- chart name, type, title, width, height, class type and series count getters
- chart create with count increment and exact property verification
- chart mutation methods exposed by the runtime (title, legend, position, size, axes, series, labels, fills/outlines, etc.)

## Runtime limitation: chart deletion

`chart.delete` is **not available through the public live co-edit Api.* object model in the deployed EuroOffice 9.3.4 runtime**.

This was measured against all public object paths available in the live runtime:

1. the `ApiChart` returned directly by `ApiWorksheet.AddChart()`;
2. the matching chart returned by `ApiWorksheet.GetAllDrawings()`;
3. the matching chart returned by `Api.GetActiveWorkbook().GetDrawingsByName(name)`.

All three expose the same chart API surface for this purpose and expose neither `Delete()` nor `Remove()`.

The current ONLYOFFICE Spreadsheet API documentation lists `ApiDrawing.Delete()` / inherited `ApiChart.Delete()`, but the deployed 9.3.4 runtime does not expose it. We therefore do not call private/internal editor methods to simulate deletion.

## Product policy

- Live co-edit remains the preferred chart transport for operations that the runtime exposes.
- `chart.delete` MUST report the live capability as unsupported rather than claim a live PASS.
- A DocBuilder implementation may be used as an explicit fallback for chart deletion.
- Any DocBuilder result MUST identify its source as `docbuilder`; it MUST NOT be reported as live co-edit.
- A DocBuilder fallback is PASS only after the returned workbook is machine-verified to contain one fewer matching chart / no chart with the requested unique name. Unmeasurable output is UNKNOWN/fail-closed.

## Probe evidence

The branch contains the reproducible probes:

- `m51-chart-delete-probe.cjs` / `m51-live-delete-probe.cjs`
- `m51-chart-object-probe.cjs` / `m51-live-chart-object-probe.cjs`

The object probe creates a uniquely named chart in one live `callCommand()`, then enumerates the exact methods of the direct chart object, worksheet drawing object, and workbook `GetDrawingsByName()` result.
