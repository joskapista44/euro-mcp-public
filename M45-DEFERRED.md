# M4.5 Excel Tables — deferred capability

Status: **DEFERRED / runtime capability**

Measured runtime: **EuroOffice DocumentServer 9.3.4-hotfix.1**.

Live-editor acceptance established:

- worksheet create/delete: PASS
- seed write + same-session live readback: PASS
- `ApiWorksheet.FormatAsTable("A1:B3")`: operation succeeds
- `ApiWorksheet.AddListObject`: unavailable in the measured runtime
- `ApiWorksheet.GetListObjects`: unavailable in the measured runtime
- therefore genuine Excel Table/ListObject identity and range cannot currently be machine-verified
- acceptance remains fail-safe: this condition is `UNKNOWN`, never manufactured as `PASS`

Upstream ONLYOFFICE DocumentServer added `ApiWorksheet.AddListObject`, `ApiWorksheet.GetListObjects`, `ApiListObject`, `ApiListColumn`, and `ApiListRow` in version 9.4.0.

## Recheck condition

When a newer EuroOffice DocumentServer based on ONLYOFFICE 9.4+ becomes available, rerun M4.5 capability/acceptance testing. Required Power User acceptance includes genuine ListObject handling and live machine-verifiable postconditions (at minimum identity/name/range; then resize and row/column operations where exposed).

Do not downgrade M4.5 to `FormatAsTable`-only PASS. Until the required runtime API becomes available, this capability is intentionally deferred and development proceeds to M4.6.
