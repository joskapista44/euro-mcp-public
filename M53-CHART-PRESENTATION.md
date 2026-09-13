# M5.3 — Live chart presentation controls

Status: **LIVE PASS / CLOSED**

Measured runtime: **EuroOffice DocumentServer 9.3.4.60** with the deployed public-getter patch.

Transport boundary:

`Playwright -> Elliot authenticated Nextcloud session -> spreadsheeteditor -> window.Asc.editor.callCommand(fn, false, callback) -> public Office Api.*`

The accepted callCommand form is the three-argument form:

`editor.callCommand(fn, false, callback)`

All acceptance evidence below comes from the live shared spreadsheet editor and uses only the public Office API at the runtime/acceptance boundary.

## Acceptance rule

M5.3 follows the project-wide fail-closed rule.

A setter call alone is not PASS. A presentation mutation is PASS only when the requested semantic state can be read back through a public Office API getter in the same live editor session and compared machine-verifiably with the requested value.

Human visual observation is not acceptance evidence.

If the matching public semantic getter is unavailable, the runtime operation remains UNKNOWN and does not perform the mutation merely to prove that the setter executed.

## Public presentation surface

For the M5.3 capabilities promoted to PASS, the live public `ApiChart` exposes the required setter/getter pairs:

- `SetLegendPos` + `GetLegendPos`
- `SetHorAxisTitle` + `GetHorAxisTitle`
- `SetVerAxisTitle` + `GetVerAxisTitle`
- `SetShowDataLabels` + `GetDataLabels`
- `ApplyChartStyle` + `GetChartStyle`

Other presentation setters also exist, but setter availability alone is not sufficient for PASS.

## Supported and machine-verified capabilities

### Legend position — LIVE PASS

Runtime operation: `chart.presentation.legend`

For a position-only request, the runtime checks both `SetLegendPos` and `GetLegendPos` before mutation.

The standalone live acceptance requested `bottom` and the public readback returned exactly `bottom`.

The earlier combined public-getter acceptance additionally verified:

- `left`
- `top`
- `right`
- `bottom`
- `none`

All tested positions received exact public readback.

Result: **PASS**

### Horizontal and vertical axis titles — LIVE PASS

Runtime operation: `chart.presentation.axisTitles`

The runtime requires the matching setter/getter pair before mutation:

- `SetHorAxisTitle` + `GetHorAxisTitle`
- `SetVerAxisTitle` + `GetVerAxisTitle`

The live getter includes the document paragraph terminator in the raw text. Acceptance therefore applies the established semantic normalization: trailing CR and LF paragraph terminators are removed before comparison.

No other title content is normalized.

The standalone acceptance requested unique horizontal and vertical titles and received exact normalized public readback for both.

Result: **PASS**

### Data labels — LIVE PASS

Runtime operation: `chart.presentation.dataLabels`

Required public pair:

- `SetShowDataLabels`
- `GetDataLabels`

The semantic readback object is compared exactly using:

- `showSerName`
- `showCatName`
- `showVal`
- `showPercent`

The standalone acceptance requested:

- `showSerName = true`
- `showCatName = true`
- `showVal = true`
- `showPercent = false`

The public getter returned the same semantic state.

The earlier combined public-getter acceptance additionally verified pie percentage labels with:

- `showSerName = false`
- `showCatName = true`
- `showVal = false`
- `showPercent = true`

Both cases passed exact machine comparison.

Result: **PASS**

### Chart style — LIVE PASS

Runtime operation: `chart.presentation.style`

Required public pair:

- `ApplyChartStyle`
- `GetChartStyle`

The deployed public getter returns the same 0-based style index used by the public setter.

The standalone acceptance exercised:

- `0 -> 0`
- `1 -> 1`
- `2 -> 2`

All three public semantic readbacks matched exactly.

Result: **PASS**

## Deliberately not claimed as PASS

### Legend font size — UNKNOWN

`SetLegendFontSize` exists, but no matching public semantic getter has been established for the project acceptance boundary.

The runtime therefore does not mutate legend font size when semantic verification is unavailable.

Result: **UNKNOWN**

This does not invalidate the PASS result of the four explicitly supported M5.3 capabilities.

The same rule applies to other presentation setters without a proven matching public semantic getter.

## Runtime fail-closed behavior

`live-chart-presentation.cjs` enforces the public getter contract.

For the supported operations:

- required setter/getter availability is checked before mutation;
- missing semantic getter returns UNKNOWN without mutation;
- wrong semantic readback returns `verification-failed`;
- exact semantic readback returns PASS.

Unit coverage proves:

- legend-position PASS;
- axis-title PASS including trailing CR/LF normalization;
- data-label PASS;
- chart-style PASS;
- missing getter -> UNKNOWN with no mutation;
- wrong getter result -> `verification-failed`;
- serialized `chartPresentationCommand` remains compatible with the live `callCommand` execution path.

## Standalone live acceptance

Acceptance file: `m53-live-acceptance.cjs`

The final standalone acceptance ran against:

- Nextcloud `https://mt-server.eu`
- authenticated user `elliot`
- workbook file ID `1231187`
- worksheet `Sheet1`
- fixture range `XFA1:XFC4`
- editor `spreadsheeteditor`
- API location `window.Asc.editor`
- source `live-coedit-editor`

The acceptance created a uniquely named bar chart and verified the presentation capability surface before exercising the semantic mutations.

Measured result:

- seed write: **PASS**
- initial chart inventory: **PASS**, count `0`
- chart creation: **PASS**, `0 -> 1`
- presentation capability inspection: **PASS**
- legend position: **PASS**
- horizontal axis title: **PASS**
- vertical axis title: **PASS**
- data labels: **PASS**
- chart style `0`: **PASS**
- chart style `1`: **PASS**
- chart style `2`: **PASS**
- cleanup delete: **PASS**, `1 -> 0`
- final chart inventory: **PASS**, count `0`
- unique chart absent after cleanup: **PASS**

Top-level acceptance result:

- milestone: `M5.3`
- source: `live-coedit-editor`
- outcome: `PASS`
- testOutcome: `PASS`
- humanObservationRequired: `false`
- exit code: `0`

The accepted cleanup path is the existing M5.1 public chart deletion route:

`ApiDrawing.Select() -> editor focus -> Playwright Delete -> public GetAllCharts() readback`

No human observation was required.

## Relationship to the combined acceptance

Before the standalone runtime path was upgraded, the combined M5.3+M5.4 public-getter acceptance independently proved the new getter semantics against the same live editor boundary.

It verified:

- legend positions `left/top/right/bottom/none`;
- horizontal and vertical axis-title readback;
- normal data-label state;
- pie percentage-label state;
- chart styles `0/1/2`.

The standalone M5.3 acceptance now additionally proves that the reusable `live-chart-presentation.cjs` runtime implementation satisfies the same public-only, fail-closed semantic contract.

## Final M5.3 result

**M5.3 — FULL LIVE PASS / CLOSED**

The supported chart presentation capabilities are machine-verifiable through the public spreadsheet API in the live editor.

M5.3 is no longer deferred because of missing presentation getters.

Individual presentation properties without a proven matching public semantic getter remain explicitly UNKNOWN and are not silently promoted to PASS.
