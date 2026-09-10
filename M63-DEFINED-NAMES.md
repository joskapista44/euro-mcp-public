# M6.3 — Defined Names

Status: **LIVE PASS**

Runtime: EuroOffice / ONLYOFFICE DocumentServer 9.3.4-hotfix.1.
Transport: `live-coedit-editor` through Playwright and the authenticated live spreadsheet editor.
Human observation required: **false**.

## Public runtime surface

The live runtime exposes `Api.AddDefName()` and `Api.GetDefName()`. A returned `ApiDefName` exposes these public methods:

- `Delete`
- `GetName`
- `GetRefersTo`
- `GetRefersToRange`
- `SetName`
- `SetRefersTo`

`Api.GetDefName(name)` throws when the name does not exist. The implementation therefore treats that specific lookup outcome as absence through a guarded lookup rather than assuming a null return.

## Implemented operations

- `definedName.inspect`
- `definedName.add`
- `definedName.setRefersTo`
- `definedName.rename`
- `definedName.delete`

All mutation operations are verified by public semantic readback. No private/internal ONLYOFFICE APIs are used. The live transport uses the deployed three-argument `editor.callCommand(fn, false, callback)` signature.

## Live acceptance

Workbook: acceptance workbook, scratch references `Sheet1!$XFD$30` and `Sheet1!$XFD$31`.

Measured sequence:

1. Add `EURO_M63_ACCEPT` referring to `=Sheet1!$XFD$30` — PASS by `GetName()` + `GetRefersTo()`.
2. Inspect after add — PASS.
3. Change reference to `=Sheet1!$XFD$31` — PASS by exact `GetRefersTo()` readback.
4. Inspect after reference change — PASS.
5. Rename to `EURO_M63_RENAMED` — PASS by exact `GetName()` + preserved `GetRefersTo()` readback.
6. Inspect after rename — PASS.
7. Delete — PASS because subsequent guarded `GetDefName()` reports absence.
8. Inspect after delete — PASS with actual state `null`.

Required operation statuses: **8/8 PASS**.
Top-level outcome: **PASS**.

Pre-clean attempts may report `not-found` when the test names are already absent; these setup results are intentionally excluded from required acceptance statuses.
