# M6.4 — Pivot Tables

Status: **LIVE PASS**

Runtime: EuroOffice / ONLYOFFICE DocumentServer 9.3.4-hotfix.1.  
Transport: `live-coedit-editor` through Playwright and the authenticated spreadsheet editor.  
Human observation required: **false**.

## Runtime-proven public surface

The deployed runtime exposes public pivot creation, lookup, field, style, source, refresh and range/data readback APIs, including `Api.InsertPivotNewWorksheet()`, `Api.GetPivotByName()`, `ApiPivotTable.AddFields()`, `AddDataField()`, `SetStyleName()`, `SetName()`, `RefreshTable()` and `GetData()`.

Worksheet cleanup is performed with public `ApiWorksheet.Delete()` and verified through `Api.GetSheets()`.

## Implemented operations

- `pivot.inspect`
- `pivot.createNewWorksheet`
- `pivot.addFields`
- `pivot.addDataField`
- `pivot.style`
- `pivot.rename`
- `pivot.refresh`

`pivot.refresh` is fail-closed. Without semantic assertions it remains UNKNOWN. With assertions, it reads public `GetData()` values before refresh and after the requested source mutation + `RefreshTable()`, and returns PASS only when all expected values match.

## Live acceptance

Source: `Sheet1!XFA60:XFC64`, with Region / Style / Price data.

Required sequence:

1. Create pivot on a new worksheet — PASS.
2. Inspect creation — PASS.
3. Add Region row field and Style column field — PASS.
4. Add Price data field — PASS by separate public inspect readback.
5. Inspect configured fields — PASS.
6. Set style `PivotStyleMedium2` — PASS.
7. Rename to `EURO_M64_RENAMED` — PASS.
8. Inspect renamed pivot — PASS.
9. Refresh semantic verification — PASS: source Price for East/A changes 10 → 110, public pivot `GetData(['East','A'])` changes 10 → 110, while control `GetData(['East','B'])` remains 30 → 30.
10. Delete the generated pivot worksheet and verify absence — PASS.

Required operation statuses: **10/10 PASS**.  
Unexpected accepted steps: **none**.  
Top-level outcome: **PASS**.

## Resolved debt

The former refresh UNKNOWN/DEFERRED state is closed. The final implementation proves the refresh effect semantically rather than treating successful `RefreshTable()` execution as evidence.

Raw `ApiRange`/parent objects are not transported through the callback; readbacks are normalized to serializable public values such as address, sheet name and scalar/matrix values.
