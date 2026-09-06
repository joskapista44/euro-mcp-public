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

## Measured runtime limitation

Direct public-object deletion is **not exposed in the deployed EuroOffice 9.3.4 spreadsheet runtime** on any path measured so far:

1. the `ApiChart` returned directly by `ApiWorksheet.AddChart()`;
2. the matching chart returned by `ApiWorksheet.GetAllDrawings()`;
3. the matching chart returned by `Api.GetActiveWorkbook().GetDrawingsByName(name)`.

Those objects expose neither `Delete()` nor `Remove()` in this deployment. The current ONLYOFFICE Spreadsheet API documentation does list `ApiDrawing.Delete()` / inherited `ApiChart.Delete()`, and marks spreadsheet `ApiDrawing.Delete()` as functionality available in paid ONLYOFFICE Docs editions. This means the correct finding is a deployed-runtime capability gap, not a general claim that ONLYOFFICE has no drawing delete API.

## Live deletion hypothesis under test

The branch now contains a fail-safe fallback that does not access private/internal ONLYOFFICE object-model fields:

1. resolve the concrete chart by its unique public `GetName()` value;
2. prefer public `Delete()` on `ApiChart` / matching drawing if the runtime exposes it;
3. otherwise require public inherited `ApiDrawing.Select()` on the matching chart/drawing;
4. select that exact chart inside `editor.callCommand(fn, false, callback)`;
5. focus the same spreadsheet editor iframe and send the editor's normal `Delete` key through Playwright;
6. re-enter `callCommand()` and verify with public `ApiWorksheet.GetAllCharts()` that the count decreased by exactly one and the unique chart name is absent.

The selection phase returns `PENDING`, never PASS. Only the post-key public readback can turn the operation into `source: "live-coedit-editor"` / `verification.status: "PASS"`.

If public `Select()` is absent, key dispatch fails, readback is unavailable, count is wrong, or the named chart remains, the operation fails closed as UNSUPPORTED / UNKNOWN / FAIL rather than claiming success.

This route uses a documented public drawing-selection API plus the editor's ordinary user-facing delete command. It intentionally avoids `AscCommon`, model internals, private drawing collections, undocumented object IDs, or mutation of internal editor state.

## Acceptance requirement

A live PASS requires one session to prove all of the following:

1. baseline `GetAllCharts()` -> N;
2. create uniquely named chart;
3. `GetAllCharts()` -> N+1;
4. delete that exact chart;
5. `GetAllCharts()` -> N;
6. unique chart name absent.

Only `source: "live-coedit-editor"` counts. DocBuilder is not a substitute for this acceptance.

## Probe evidence

The branch contains the reproducible probes:

- `m51-chart-delete-probe.cjs` / `m51-live-delete-probe.cjs`
- `m51-chart-object-probe.cjs` / `m51-live-chart-object-probe.cjs`

The object probe creates a uniquely named chart in one live `callCommand()`, then enumerates the methods of the direct chart object, worksheet drawing object, and workbook `GetDrawingsByName()` result.
