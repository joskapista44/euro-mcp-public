# M5.2 — Live chart series management

M5.2 extends the M5.1 live chart path with machine-verifiable series operations in the current ONLYOFFICE spreadsheet editor.

## Source boundary

All PASS results must come from:

`Playwright -> spreadsheeteditor -> window.Asc.editor.callCommand(fn, false, callback) -> public Office Api.*`

`source` must be `live-coedit-editor`. No saved-file, OOXML, LibreOffice or DocBuilder path may claim live PASS.

## Supported operations

- `chart.series.inspect`
  - `ApiChart.GetAllSeries()`
  - per-series `GetChartType()` and `GetClassType()` when exposed
- `chart.series.add`
  - `ApiChart.AddSeria(nameRange, valuesRange, xValuesRange?)`
  - PASS requires live series count `N -> N+1`
- `chart.series.remove`
  - `ApiChart.RemoveSeria(index)`
  - PASS requires live series count `N -> N-1`
- `chart.series.changeType`
  - `ApiChartSeries.ChangeChartType(type)`
  - PASS requires same-session `GetChartType()` to equal the requested type after mutation

## Why this scope

The public Spreadsheet `ApiChartSeries` surface exposes `ChangeChartType`, `GetChartType` and `GetClassType`. Series creation/removal lives on `ApiChart` and can be verified independently by `GetAllSeries()` count readback. These operations therefore satisfy the project fail-closed verification contract without relying on visual observation.

Presentation setters that lack a matching public getter are not accepted as M5.2 PASS merely because a setter returns successfully. They can be considered in a later milestone if a stable machine-readable postcondition is available.

## Acceptance

The live acceptance uses a temporary uniquely named combo chart and requires:

1. exact seed-data live readback;
2. chart create with two series;
3. initial series inspect count = 2;
4. series 0 type change with exact `GetChartType()` readback;
5. add a third series and read back count = 3;
6. remove that third series and read back count = 2;
7. final series inspect preserves the changed series type;
8. delete the temporary chart using the already accepted M5.1 live delete path;
9. final chart inventory returns to the baseline count and the unique chart name is absent.

UNKNOWN / UNSUPPORTED is never PASS. Human visual observation is not required and is not sufficient for PASS.
