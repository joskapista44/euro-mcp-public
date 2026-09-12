# Excel Power User capability gap register

Canonical EURO-MCP register for Excel Power User capabilities that are not currently machine-verifiable through the deployed EuroOffice / ONLYOFFICE public API.

Last live recheck: 2026-09-12
Deployed EuroOffice DocumentServer: 9.3.4.60

## Acceptance rule

EURO-MCP is fail-closed. A setter call alone is not PASS. A mutation is PASS only when its semantic postcondition can be machine-read back through a stable public Office API in the live editor session. Human visual observation is not acceptance evidence. Private/internal ONLYOFFICE APIs are not verification paths.

Transport: MCP -> Playwright -> authenticated Nextcloud spreadsheet editor -> editor.callCommand(fn, false, callback) -> public Office Api.

## Current gaps

### M4.5 Excel Tables / ListObject — DEFERRED

Fresh 9.3.4.60 acceptance confirms that FormatAsTable executes, but ApiWorksheet.AddListObject and ApiWorksheet.GetListObjects are unavailable. Genuine table identity, name and range therefore cannot be semantically verified. Do not treat FormatAsTable alone as Table/ListObject PASS.

Recheck when EuroOffice is based on ONLYOFFICE 9.4 or newer, where the upstream ListObject API became available.

### M5.3 Advanced chart presentation — DEFERRED

Fresh 9.3.4.60 runtime probe confirms:

- legend: SetLegendPos exists; public GetLegendPos absent;
- axis titles: SetHorAxisTitle and SetVerAxisTitle exist; matching semantic getters absent;
- data labels: SetShowDataLabels and SetShowPointDataLabel exist; matching state getter absent;
- chart style: ApplyChartStyle exists; chart-style getter absent.

These mutations remain DEFERRED until public semantic readback exists.

### M5.4 Chart data and object management — PARTIAL PASS / DEFERRED

Verified PASS:

- rename: SetName + GetName;
- resize: SetSize + GetWidth/GetHeight.

Open gaps on 9.3.4.60:

- position — DEFERRED: SetPosition exists; GetPosition, GetPosX and GetPosY absent;
- series name/values/X-values/category formula — DEFERRED: SetSeriaName, SetSeriaValues, SetSeriaXValues and SetCatFormula exist; matching public semantic getters absent;
- copy/duplicate — UNSUPPORTED: chart Copy absent and worksheet AddDrawing absent.

The public series wrapper exposes ChangeChartType, GetChartType and GetClassType. private_GetSeries is internal and forbidden for acceptance.

### M7 Protection — CORE LIVE PASS WITH DEFERRED EXTENSIONS

Core protected-range ACL is LIVE PASS for create, inspect, rename, add-user, delete-user and persistence.

Remaining gaps:

- classic whole-sheet password protection — UNSUPPORTED / no deployed public API measured;
- SetRange — DEFERRED: setter exists, semantic range getter not proven;
- SetAnyoneType — DEFERRED: setter exists, semantic ACL getter/readback not proven;
- delete protected range — UNSUPPORTED: no measured public ApiProtectedRange.Delete.

## Completed advanced capabilities

These are not gaps:

- M6.1 Sort & Filter — LIVE PASS
- M6.2 Data Validation — LIVE PASS
- M6.3 Defined Names — LIVE PASS
- M6.4 Pivot Tables — LIVE PASS, including semantic refresh readback
- M7 Protected Range ACL core — LIVE PASS
- chart rename — PASS
- chart resize — PASS
- Freeze Panes — implemented earlier

## Status meanings

DEFERRED means an operation is present or partly present but EURO-MCP cannot prove its semantic result through the deployed public API. UNSUPPORTED means the required public operation itself has not been measured on the deployed runtime. Neither status may be converted to PASS by visual inspection, private APIs, or an unrelated formatting operation.

## Maintenance rule

This file is part of the MCP documentation. Update it whenever the DocumentServer runtime changes, a capability probe discovers a new public operation/readback, a deferred capability becomes machine-verifiable, or a new Excel Power User gap is found.

After all implementable Power User capabilities are closed or recorded here as runtime-deferred/unsupported, run a final cross-capability Excel Power User integrated acceptance. Documented runtime blockers should be reported separately rather than treated as implementation failures.
