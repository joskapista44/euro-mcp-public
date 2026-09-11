# M6.1 — Sort & Filter

Status: **LIVE PASS**

Runtime: EuroOffice / ONLYOFFICE DocumentServer 9.3.4-hotfix.1.  
Transport: `live-coedit-editor` through Playwright and the authenticated spreadsheet editor.  
Human observation required: **false**.

## Implemented operations

- `sort.apply`
- `filter.inspect`
- `filter.enable`
- `filter.set`
- `filter.reapply`
- `filter.clear`

Sort uses the public `ApiRange.SetSort()` surface and requires an explicit `keyRange`. Filter operations use public `ApiRange.SetAutoFilter()` and `ApiAutoFilter` readbacks. Unknown or unverifiable states fail closed.

## Verification model

Sort reads the range before and after mutation with public `ApiRange.GetValue()`. Filter mutations are checked through `ApiWorksheet.GetAutoFilter()`, its range, filter mode and readable filter criteria. Successful transport alone is not PASS.

## Live acceptance

Scratch source: `Sheet1!XFA1:XFC5`.

The acceptance seeds an unsorted data set, sorts by the Name column, reads the matrix back and requires the exact order `Name, Alpha, Bravo, Charlie, Delta`. It then enables AutoFilter, applies a Group=`A` criterion, inspects the live state, reapplies the filter and clears it. All accepted operations require public semantic readback.

Top-level outcome: **PASS**.

## Debt closure

`test_live_sort_filter.cjs` was aligned with the deployed contract: `sort.apply` now tests mandatory `keyRange`, `filter.set` replaces the obsolete mock operation, and reapply/clear assertions match the current public API semantics.
