# M5.4 — Chart data and object management

## Status

**LIVE PASS** on deployed EuroOffice DocumentServer 9.3.4.60 with one documented runtime limitation:

- chart copy/duplicate: **UNSUPPORTED**

Live transport:

`Playwright → authenticated Nextcloud session → ONLYOFFICE spreadsheeteditor → window.Asc.editor.callCommand()`

Source: `live-coedit-editor`

Human observation required: **false**

EURO-MCP remains fail-closed: a setter call alone is not PASS. Every supported mutation below is accepted only through machine-verifiable semantic readback using the public spreadsheet Office API.

## Verified PASS capabilities

### Chart rename

`SetName()` followed by exact public `GetName()` readback.

Live acceptance: **PASS**

### Chart resize

`SetSize()` followed by exact public `GetWidth()` / `GetHeight()` readback.

Live acceptance: **PASS**

### Chart position

`SetPosition(fromCol, colOffset, fromRow, rowOffset)` followed by public `GetPosition()` semantic readback.

The accepted public getter returns:

- `fromCol`
- `colOffset`
- `fromRow`
- `rowOffset`

The combined getter acceptance exercised three distinct positions with exact semantic readback.

The standalone M5.4 acceptance also exercised:

`{ fromCol: 7, colOffset: 180000, fromRow: 9, rowOffset: 540000 }`

and received the exact same public readback.

Live acceptance: **PASS**

### Series name

`SetSeriaName()` followed by public `ApiChartSeries.GetName()` readback.

The runtime represents a literal series name as an Excel string formula such as:

`="EURO Series ..."`

Acceptance decodes that public representation back to its semantic string value before exact comparison.

Live acceptance: **PASS**

### Series values

`SetSeriaValues()` followed by public `ApiChartSeries.GetValues()` readback.

The runtime may return the range formula with one leading `=`. Acceptance removes exactly one leading `=` before semantic comparison.

Live acceptance: **PASS**

### Category formula

`SetCatFormula()` followed by public `ApiChartSeries.GetCatFormula()` readback.

The same one-leading-`=` formula normalization is used.

Live acceptance: **PASS**

### Scatter X-values

Spreadsheet range-based X-value mutation is verified on a real scatter chart:

`SetSeriaXValues(range, seriesIndex)` followed by public `ApiChartSeries.GetXValues()` readback.

This is intentionally tested on a separate scatter fixture rather than forcing X-value semantics onto the bar-chart fixture.

Live acceptance: **PASS**

### Object inspection

Public object inspection verifies:

- name
- width
- height
- position setter/readback availability
- series setter/readback availability
- copy capability availability

Live acceptance: **PASS**

## Runtime limitation — chart copy / duplicate

### Copy

Chart copy/duplicate remains **UNSUPPORTED** on the deployed spreadsheet runtime.

Live public-only probing established that:

- the live spreadsheet `ApiChart` object does not expose a public `Copy()` method;
- the public worksheet surface does not expose `AddDrawing()` or another measured generic drawing-attach path suitable for attaching a copied chart.

The standalone M5.4 acceptance therefore returns:

`copyDuplicate: UNSUPPORTED`

This is a measured runtime limitation, not a semantic-readback failure and not a DEFERRED implementation.

Copy methods found in other editor contexts inside the combined DocumentServer SDK, and internal spreadsheet copy/attach mechanisms, do not satisfy the spreadsheet public-API acceptance boundary.

No fake copy implementation is reported as PASS.

## Public semantic normalization rules

The accepted public API exposes some values in formula/document representations rather than the logical input representation.

Machine verification therefore uses these narrow semantic normalizations:

1. Series name:
   - `="foo"` → `foo`
   - doubled Excel quotes inside the literal are decoded.

2. Series formulas:
   - exactly one leading `=` is removed before comparison.

These normalizations operate only on public getter output. No private/internal chart state is used for acceptance.

## Live acceptance

The final standalone M5.4 acceptance against the shared live workbook produced:

- seed bar data: PASS
- seed scatter data: PASS
- initial chart inventory: PASS (`0`)
- create bar chart: PASS (`0 → 1`)
- create scatter chart: PASS (`1 → 2`)
- object capability inspection: PASS
- rename: PASS with exact public readback
- resize: PASS with exact public readback
- position: PASS with exact public `GetPosition()` readback
- series name: PASS with normalized public `GetName()` readback
- series values: PASS with normalized public `GetValues()` readback
- category formula: PASS with normalized public `GetCatFormula()` readback
- scatter X-values: PASS with normalized public `GetXValues()` readback
- copy/duplicate: UNSUPPORTED
- scatter cleanup delete: PASS via `public-select+editor-delete-key` (`2 → 1`)
- main cleanup delete: PASS via `public-select+editor-delete-key` (`1 → 0`)
- final chart inventory: PASS (`0`)

Overall standalone acceptance result: **PASS**

Exit code: **0**

The earlier combined M5.3+M5.4 public-getter acceptance also produced **PASS** for position and all four series-data semantic readbacks.

## Acceptance boundary

All PASS claims above use:

`MCP/runtime → Playwright → authenticated live spreadsheet editor → editor.callCommand(fn, false, callback) → public Office Api`

Private helpers may be used inside the DocumentServer implementation of a public getter, but acceptance itself calls only the public API.

Human visual observation is not acceptance evidence.

Internal spreadsheet chart state is not acceptance evidence.

## Final M5.4 result

**M5.4 — LIVE PASS**

Supported and machine-verified:

- rename
- resize
- position
- series name
- series values
- category formula
- scatter X-values
- object inspection

Documented deployed-runtime limitation:

- copy/duplicate — **UNSUPPORTED**

The unsupported copy/duplicate operation remains explicitly visible rather than being converted into PASS or hidden behind the milestone result.
