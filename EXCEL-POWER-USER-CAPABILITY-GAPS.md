# Excel Power User capability gap register

Canonical EURO-MCP register for Excel Power User capabilities that are not currently machine-verifiable through the deployed EuroOffice / ONLYOFFICE public API.

Last live recheck: 2026-09-17
Deployed EuroOffice DocumentServer: 9.3.4.60

## Acceptance rule

EURO-MCP is fail-closed. A setter call alone is not PASS. A mutation is PASS only when its semantic postcondition can be machine-read back through a stable public Office API in the live editor session. Human visual observation is not acceptance evidence. Private/internal ONLYOFFICE APIs are not verification paths.

Transport: MCP -> Playwright -> authenticated Nextcloud spreadsheet editor -> editor.callCommand(fn, false, callback) -> public Office Api.

## Current gaps

### Print layout / verified one-page output — DEVELOPMENT BACKLOG

The current `office_xlsx_batch` surface can compose cell content, formulas, formatting, fixed row/column dimensions, merges and drawing objects, but it does not inspect rendered print pagination or expose a verified print-layout final state.

Required W0.12 work:

- probe public live setters and getters for print area, paper size, portrait/landscape orientation, margins, scaling and fit-to-pages width/height;
- expose only settings whose exact post-state can be read back in the same editor session;
- determine whether the public runtime exposes calculated page breaks or rendered page count;
- if pagination itself is not publicly measurable, report only `fit-to-one-page settings live-verified`, never claim that the rendered output is one page;
- add persisted retry and a cross-sheet print-layout acceptance without Chrome UI automation, fixed delay, downloaded XLSX or OOXML verification.

Known legacy probes are insufficient: some setters were inert or unavailable on the deployed runtime, and setter dispatch alone is not acceptance evidence. Status: DEVELOPMENT BACKLOG / runtime re-probe required.

### M4.5 Excel Tables / ListObject — DEFERRED

Fresh 9.3.4.60 acceptance confirms that FormatAsTable executes, but ApiWorksheet.AddListObject and ApiWorksheet.GetListObjects are unavailable. Genuine table identity, name and range therefore cannot be semantically verified. Do not treat FormatAsTable alone as Table/ListObject PASS.

Recheck when EuroOffice is based on ONLYOFFICE 9.4 or newer, where the upstream ListObject API became available.

### M5.3 Advanced chart presentation — LIVE PASS

The deployed 9.3.4.60 DocumentServer has been extended with public semantic getters required for fail-closed acceptance.

Live public-only acceptance now verifies:

- legend position: SetLegendPos + GetLegendPos;
- horizontal and vertical axis titles: setters + public semantic title getters;
- data-label state: setter + public GetDataLabels readback;
- chart style: ApplyChartStyle + public GetChartStyle readback.

M5.3 is no longer a capability gap.

### M5.4 Chart data and object management — LIVE PASS WITH DOCUMENTED RUNTIME LIMITATION

Verified public-only LIVE PASS:

- rename: SetName + GetName;
- resize: SetSize + GetWidth/GetHeight;
- position: SetPosition + GetPosition;
- series name: SetSeriaName + ApiChartSeries.GetName;
- series values: SetSeriaValues + ApiChartSeries.GetValues;
- category formula: SetCatFormula + ApiChartSeries.GetCatFormula;
- scatter X-values: SetSeriaXValues + ApiChartSeries.GetXValues.

The final standalone M5.4 live acceptance returned PASS with exit code 0 and cleaned both test charts back to a final chart count of zero.

Documented deployed-runtime limitation:

- copy/duplicate — UNSUPPORTED: the live spreadsheet ApiChart object exposes no public Copy(), and the public worksheet surface exposes no AddDrawing or other measured generic drawing-attach path suitable for a copied chart.

This copy limitation is not a DEFERRED semantic-readback gap. No public spreadsheet copy operation is available to implement and verify on the deployed runtime.

Acceptance uses only public Office API readback. Internal/private series state is forbidden as acceptance evidence.

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
- M5.3 Advanced chart presentation — LIVE PASS
- M5.4 Chart data and object management — LIVE PASS, with copy/duplicate documented as runtime-UNSUPPORTED
- Freeze Panes — implemented earlier

## Status meanings

DEFERRED means an operation is present or partly present but EURO-MCP cannot prove its semantic result through the deployed public API. UNSUPPORTED means the required public operation itself has not been measured on the deployed runtime. Neither status may be converted to PASS by visual inspection, private APIs, or an unrelated formatting operation.

## Maintenance rule

This file is part of the MCP documentation. Update it whenever the DocumentServer runtime changes, a capability probe discovers a new public operation/readback, a deferred capability becomes machine-verifiable, or a new Excel Power User gap is found.

After all implementable Power User capabilities are closed or recorded here as runtime-deferred/unsupported, run a final cross-capability Excel Power User integrated acceptance. Documented runtime blockers should be reported separately rather than treated as implementation failures.

## W0.12 migration closure

The final real-MCP cross-capability gate passed on 2026-09-17. One `office_xlsx_batch` invocation applied and semantically verified 13 compatible Power User goals in one persistent editor session with one save barrier; the persisted retry reopened once and proved the same final state with 0 writes and no barrier. Runtime-deferred and unsupported entries above remain explicit exclusions, not failed acceptance items.
