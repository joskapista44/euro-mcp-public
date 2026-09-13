# EuroOffice DocumentServer 9.3.4.60 chart public getter patch

This directory contains the versioned, fail-closed patch for the accepted
EuroOffice DocumentServer 9.3.4.60 `sdk-all.js` chart getter extension.

It adds and publicly exports these ten getters:

- `ApiChart.GetLegendPos`
- `ApiDrawing.GetPosition`
- `ApiChart.GetHorAxisTitle`
- `ApiChart.GetVerAxisTitle`
- `ApiChart.GetDataLabels`
- `ApiChart.GetChartStyle`
- `ApiChartSeries.GetName`
- `ApiChartSeries.GetValues`
- `ApiChartSeries.GetXValues`
- `ApiChartSeries.GetCatFormula`

Known input:

`2e5059f40b8e4c669b1c24d1724423aa8de68a957181a71cd616bb9bf742cbce`

Accepted patched output:

`1ea9602263e0408452608a1862b31b72759a143f6df117f8f6bd7d41faf0cc03`

The patcher is intentionally fail-closed.

States:

- `PATCHABLE`: exact known original bundle.
- `ALREADY_PATCHED`: exact accepted patched bundle.
- `NATIVE_OR_UNKNOWN_SUPPORT`: unknown bundle containing all ten public getter
  signatures. It is not modified; live semantic acceptance is required.
- `INCOMPATIBLE`: unknown or structurally unexpected bundle. It is not modified.

An unknown future EuroOffice release must never receive this 9.3.4.60 patch
blindly.

The manifest contains nine deterministic replacements:

- six getter implementation regions;
- three public/export registry regions.

The reference patched bundle differs from the original by exactly 3750 bytes:
1944 bytes of getter implementations plus 1806 bytes of public/export wiring.

The full 20 MB `sdk-all.js` bundle is deliberately not stored in Git.
