# M5.1 chart runtime capability — EuroOffice 9.3.4

Measured live on 2026-09-06 through the production co-edit transport:

`MCP -> Playwright -> Elliot authenticated Nextcloud session -> ONLYOFFICE spreadsheeteditor -> editor.callCommand() -> Api.*`

## M5.1 status

**FINAL LIVE PASS: create / inspect / modify / delete.**

The acceptance is machine-read-back verified, uses `source: "live-coedit-editor"`, and requires no human observation.

## Live co-edit capabilities

PASS / machine-read-back verified:

- `ApiWorksheet.AddChart`
- `ApiWorksheet.GetAllCharts`
- `ApiWorksheet.GetAllDrawings`
- chart name, type, title, width, height, class type and series count getters
- chart create with count increment and exact property verification
- chart mutation methods exposed by the runtime (title, legend, position, size, axes, series, labels, fills/outlines, etc.)
- chart delete through the measured public-selection/editor-command route described below

## Measured direct-delete runtime limitation

Direct public-object deletion is **not exposed in the deployed EuroOffice 9.3.4 spreadsheet runtime** on the measured object paths:

1. the `ApiChart` returned directly by `ApiWorksheet.AddChart()`;
2. the matching chart returned by `ApiWorksheet.GetAllDrawings()`;
3. the matching chart returned by `Api.GetActiveWorkbook().GetDrawingsByName(name)`.

Those measured objects expose neither `Delete()` nor `Remove()` in this deployment. The current ONLYOFFICE Spreadsheet API documentation does list `ApiDrawing.Delete()` / inherited `ApiChart.Delete()`, and marks spreadsheet `ApiDrawing.Delete()` as functionality available in paid ONLYOFFICE Docs editions. This remains a deployed-runtime capability finding, not a general claim that ONLYOFFICE has no drawing delete API.

## Live chart deletion — measured PASS

The branch implements and has now live-verified a fail-safe fallback that does not access private/internal ONLYOFFICE object-model fields:

1. resolve the concrete chart by its unique public `GetName()` value;
2. prefer public `Delete()` on `ApiChart` / matching drawing if a future runtime exposes it;
3. otherwise require public inherited `ApiDrawing.Select()` on the matching chart/drawing;
4. select that exact chart inside `editor.callCommand(fn, false, callback)`;
5. focus the same spreadsheet editor iframe and send the editor's normal `Delete` key through Playwright;
6. re-enter `callCommand()` and verify with public `ApiWorksheet.GetAllCharts()` that the count decreased by exactly one and the unique chart name is absent.

The selection phase is never treated as deletion success. Only the post-key public readback can produce `source: "live-coedit-editor"` / `verification.status: "PASS"`.

This route uses a documented public drawing-selection API plus the editor's ordinary user-facing delete command. It intentionally avoids `AscCommon`, model internals, private drawing collections, undocumented object IDs, or mutation of internal editor state.

If public `Select()` is absent, key dispatch fails, readback is unavailable, count is wrong, or the named chart remains, the operation fails closed as UNSUPPORTED / UNKNOWN / FAIL rather than claiming success.

## Live acceptance evidence — 2026-09-06

Runtime: EuroOffice DocumentServer `9.3.4-hotfix.1`.

Acceptance file: `Flotta közös/elliot/EURO-MCP/debug.xlsx`.

Measured in one live authenticated co-edit session:

1. seed write: PASS;
2. exact seed readback: PASS;
3. baseline `GetAllCharts()`: `0`;
4. create uniquely named chart `M51_mtqc9vw7`: count `0 -> 1`, PASS;
5. create property verification: bar chart, one series, expected name/title/size, PASS;
6. modify title and size `3600000 x 2200000 -> 4000000 x 2400000`: PASS;
7. separate post-modify inspect: PASS;
8. delete route: `public-select+editor-delete-key`;
9. delete machine readback: count `1 -> 0`, `nameAbsent: true`, PASS;
10. separate post-delete inspect: count `0`, chart list empty, PASS;
11. final delete acceptance: `matches: true`;
12. top-level `outcome: "PASS"`, `testOutcome: "PASS"`, `source: "live-coedit-editor"`, `humanObservationRequired: false`.

This satisfies the M5.1 live chart-delete acceptance contract: baseline N, create N+1, uniquely identified chart, delete, return to N, and unique name absent.

## Acceptance contract

A live PASS requires one session to prove all of the following:

1. baseline `GetAllCharts()` -> N;
2. create uniquely named chart;
3. `GetAllCharts()` -> N+1;
4. delete that exact chart;
5. `GetAllCharts()` -> N;
6. unique chart name absent.

Only `source: "live-coedit-editor"` counts. DocBuilder is not a substitute for this acceptance. UNKNOWN / UNSUPPORTED remains fail-safe and never counts as PASS.

## Probe evidence

The branch contains the reproducible probes:

- `m51-chart-delete-probe.cjs` / `m51-live-delete-probe.cjs`
- `m51-chart-object-probe.cjs` / `m51-live-chart-object-probe.cjs`

The object probe creates a uniquely named chart in one live `callCommand()`, then enumerates the methods of the direct chart object, worksheet drawing object, and workbook `GetDrawingsByName()` result.
