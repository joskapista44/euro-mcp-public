# M7 — Protected Range ACL

Status: **LIVE PASS** on the deployed EURO-OFFICE spreadsheet editor.

## Scope

M7 implements the deployed public ONLYOFFICE Protected Range ACL surface through the existing live co-edit transport:

`MCP -> Playwright -> authenticated Nextcloud editor -> window.Asc.editor.callCommand(fn, false, callback) -> public Office Api`

The implementation intentionally does not use private/internal APIs.

## Verified operations

The following operations have machine-verifiable semantic live readback:

- create protected range: `ApiWorksheet.AddProtectedRange` + `GetProtectedRange`
- inspect protected range: `GetProtectedRange`, `GetAllProtectedRanges`
- rename protected range: `ApiProtectedRange.SetTitle`; the deployed editor can swallow immediate callback/readback responses after this mutation, so semantic verification is obtained by a subsequent public lookup/use of the new title. The integrated acceptance requires a successful `GetProtectedRange(newTitle)` path followed by exact public user ACL readback.
- add named user ACL: `AddUser` + exact `GetUser` (`id`, `name`, `type`)
- delete named user ACL: `DeleteUser` + delayed separate public `GetUser` / `GetAllUsers` readback proving absence
- persistence across editor reload: protected range and named user ACL survive reload, machine-proven by `m7-protected-range-persistence-probe.cjs`

The accepted transport signature remains exactly:

`window.Asc.editor.callCommand(fn, false, callback)`

## Integrated live acceptance

`m7-live-acceptance.cjs` executes:

create -> inspect -> rename -> add-user -> inspect -> delete-user -> inspect

Live result: **PASS** for every required step, with `humanObservationRequired: false`.

## Deferred / unsupported-for-verified-operation

These are not counted as M7 failures because the deployed public API does not currently provide the semantic surface needed by the project's fail-closed acceptance rules:

- classic whole-sheet password protection: no measured deployed public API
- `SetRange`: mutation method exists, but no public semantic range getter has been proven
- `SetAnyoneType`: mutation method exists, but no public semantic getter/readback has been proven
- deleting a protected range: no measured public `ApiProtectedRange.Delete` method

Private `_checkProtection` is explicitly forbidden and is not used.

## Runtime caveat

`SetTitle` has a deployed-runtime callback anomaly: after the mutation, immediate and even repeated standalone `callCommand` lookup callbacks may be swallowed. The rename itself is nevertheless semantically verified because the next public operation resolves the protected range under the new title and completes an exact public ACL mutation/readback. The acceptance records this verification path explicitly rather than manufacturing a callback-based PASS.

## Fixture debt

The deployed API exposes no measured public protected-range delete operation. Consequently historical live probes/acceptance runs leave protected-range fixtures in `debug.xlsx`. Future regression acceptance should reuse/normalize a fixed fixture rather than creating an unbounded sequence of ranges.
