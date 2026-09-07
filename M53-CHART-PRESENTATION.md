# M5.3 — Live chart presentation controls

Status: **DEFERRED / runtime verification capability**

Measured runtime: **EuroOffice DocumentServer 9.3.4-hotfix.1**.

Transport boundary:

`Playwright -> Elliot authenticated Nextcloud session -> spreadsheeteditor -> window.Asc.editor.callCommand(fn, false, callback) -> public Office Api.*`

## Measured live setter surface

The deployed `ApiChart` exposes public setters for presentation including:

- legend: `SetLegendPos`, `SetLegendFontSize`, fill/outline;
- axes: horizontal/vertical titles, label font size, orientation, tick marks and number format;
- gridlines: major/minor horizontal and vertical gridline setters;
- labels: `SetShowDataLabels`, `SetShowPointDataLabel`;
- style/fill/outline: `ApplyChartStyle`, plot/title/series/data-point/marker fill and outline setters.

## Verification limitation

The same deployed public chart object does **not** expose semantic getters for the presentation state needed to verify those writes. In particular the runtime probe found no public `GetLegendPos`, axis-title getter, data-label-state getter, or chart-style getter.

A successful setter call is therefore not sufficient for project PASS. M5.3 deliberately returns `verification.status: "UNKNOWN"` after such mutations unless a future runtime supplies a matching public getter. Human visual observation is not accepted as a substitute.

`chart.presentation.inspect` is PASS-capable because method availability itself is machine-readable. The mutation operations are implemented fail-closed:

- `chart.presentation.legend`
- `chart.presentation.axisTitles`
- `chart.presentation.dataLabels`
- `chart.presentation.style`

If a future runtime exposes `GetLegendPos`, the legend-position path already upgrades to exact readback PASS when the getter matches the requested value. Other controls remain UNKNOWN until equivalent public semantic readback exists.

## Acceptance contract

The live acceptance creates a uniquely named chart, verifies the live presentation capability matrix, exercises the setter paths, requires every unverifiable mutation to remain UNKNOWN rather than false PASS, deletes the chart through the accepted M5.1 public-select/editor-delete route, and verifies the chart inventory returns to baseline.

The expected top-level result on 9.3.4-hotfix.1 is **DEFERRED**, not PASS. This is a runtime capability result, not a transport failure.

## Recheck condition

Re-run M5.3 after upgrading the EuroOffice/ONLYOFFICE runtime. Promote each presentation operation to PASS only when its requested state can be read back through a stable public semantic API in the same live session.
