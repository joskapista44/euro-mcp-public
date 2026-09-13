# M6 — Excel Power User data operations closure

Status: **CLOSED / LIVE PASS**

M6 groups four advanced spreadsheet data-operation capabilities that were proven against the authenticated live EuroOffice / ONLYOFFICE editor. The common transport is Playwright → live spreadsheet editor → three-argument `editor.callCommand(fn, false, callback)` → public Office API.

## Closure matrix

| Milestone | Capability | Result |
|---|---|---|
| M6.1 | Sort & Filter | **LIVE PASS** |
| M6.2 | Data Validation | **LIVE PASS** |
| M6.3 | Defined Names | **LIVE PASS** |
| M6.4 | Pivot Tables | **LIVE PASS** |

## Acceptance rule

M6 follows the project-wide fail-closed rule: a mutation is not PASS merely because the command executes. PASS requires machine-verifiable semantic readback through a public ONLYOFFICE API. Human observation is not acceptance evidence.

All four M6 sub-milestones have live machine-verifiable PASS evidence and require no private/internal ONLYOFFICE APIs.

## Debt status at closure

- M6.1 stale unit test contract: **closed**; mocks and assertions aligned with the deployed sort/filter contract.
- M6.1 documentation: **closed** (`M61-SORT-FILTER.md`).
- M6.2 documentation: **closed** (`M62-DATA-VALIDATION.md`).
- M6.3 documentation: **already complete and retained** (`M63-DEFINED-NAMES.md`).
- M6.4 refresh verification: **closed** with public `ApiPivotTable.GetData()` semantic before/after readback.
- M6.4 documentation: **closed** (`M64-PIVOT-TABLES.md`).

## Explicitly outside M6 closure

The following are not hidden M6 blockers and remain separate future/deferred capability work:

- Excel Tables / ListObject support, previously deferred because the deployed runtime lacked a usable real table/ListObject API.
- M5.3/M5.4 advanced chart work was outside the original M6 closure scope; it has since been resolved separately: M5.3 is FULL LIVE PASS / CLOSED, and M5.4 is LIVE PASS for the supported chart data/object capabilities, with chart copy/duplicate explicitly documented as runtime-UNSUPPORTED on EuroOffice DocumentServer 9.3.4.60.
- Worksheet/range Protection, to be assessed as a later capability milestone.
- Final cross-capability Excel Power User integrated acceptance, to be performed after the remaining planned capability work.

Freeze Panes was already implemented earlier and is not reopened by M6.

## Repository state

M6 is closed on `feature/m6-power-user-next`. Closure does **not** itself authorize or perform a merge to `main`; branch integration remains a separate decision/gate.
