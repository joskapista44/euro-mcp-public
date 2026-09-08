# M5.4 — Chart data and object management

## Status

**DEFERRED** on deployed EuroOffice DocumentServer 9.3.4-hotfix.1.

Live transport:

`Playwright → authenticated Nextcloud session → ONLYOFFICE spreadsheeteditor → window.Asc.editor.callCommand()`

Source: `live-coedit-editor`

Human observation required: **false**

## Verified PASS capabilities

### Chart rename

`SetName()` followed by exact public `GetName()` readback.

Live acceptance: **PASS**

### Chart resize

`SetSize()` followed by exact public `GetWidth()` / `GetHeight()` readback.

Live acceptance: **PASS**

### Object inspection

Public name, width, height and runtime capability inspection.

Live acceptance: **PASS**

## Fail-closed deferred capabilities

### Position

`SetPosition()` exists, but the deployed runtime exposes no public semantic position getter.

Result: **UNKNOWN**

No mutation is attempted.

### Series data

The deployed runtime exposes:

- `SetSeriaName()`
- `SetSeriaValues()`
- `SetSeriaXValues()`
- `SetCatFormula()`

but exposes no corresponding public semantic getters suitable for machine verification.

Results: **UNKNOWN**

No mutation is attempted.

`private_GetSeries` is explicitly not used for acceptance.

### Copy

The deployed runtime exposes neither a usable public `Copy()` path nor `Worksheet.AddDrawing()`.

Result: **UNSUPPORTED**

## Live acceptance

Measured against the shared live workbook:

- seed: PASS
- initial chart inventory: PASS
- create chart: PASS (`0 → 1`)
- object capability inspection: PASS
- rename: PASS with exact readback
- resize: PASS with exact readback
- position: UNKNOWN, no mutation
- series name: UNKNOWN, no mutation
- series values: UNKNOWN, no mutation
- series X values: UNKNOWN, no mutation
- category formula: UNKNOWN, no mutation
- copy: UNSUPPORTED
- cleanup delete: PASS via `public-select+editor-delete-key` (`1 → 0`)
- final chart inventory: PASS (`0`)

Overall acceptance result: **DEFERRED**

Reason:

> rename and resize have exact public readback PASS; position and series-data setters lack public semantic readback, while copy/AddDrawing is unavailable on deployed runtime

This is intentional fail-closed behavior. No capability without machine-verifiable semantic readback is reported as PASS.
