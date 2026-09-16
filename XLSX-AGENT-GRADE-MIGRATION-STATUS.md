# XLSX agent-grade persistent execution migration status

Audit date: 2026-09-16
Branch: `feature/xlsx-persistent-live-session`
Runtime acceptance workbook: FILE_ID 1231187

## Acceptance model

Agent-grade means more than an Office API mutation succeeding. For a supported task the executor must use one persistent live editor session, establish fresh LIVE_READ identity at mutation boundaries, fail closed on ambiguity or unverifiable state, require semantic LIVE_VERIFY readback, perform whole-task verification, persist only after verified writes, and make a persisted retry either a proven no-op or a fail-closed conflict.

`IMPLEMENTED`, `STATIC PASS`, primitive LIVE PASS and whole-task TRUE LIVE PASS are distinct states.

## Canonical executor

`executeAgentTaskInPersistentSession` is the canonical persistent executor for the generic task vocabulary currently implemented in `xlsx-agent-task.cjs`:

- `create_sheet`
- `write_range`
- `copy_sheet`
- `rename_sheet`
- `delete_sheet`

The canonical executor now performs persisted copy-chain final-state classification before generic redispatch. This covers copy -> rename and copy -> downstream write(s) -> rename final states without requiring the former dedicated copy-chain wrapper.

## TRUE LIVE PASS

The following paths have direct runtime acceptance on the persistent-session branch:

- create_sheet + write_range, including formula semantics and persistence barrier;
- rename_sheet;
- delete_sheet and persisted retry;
- copy_sheet under the current narrow content/formula usedRange contract and persisted retry;
- copy -> write -> rename persisted retry;
- copy -> expanded downstream write -> rename persisted retry;
- copy -> rename persisted retry;
- canonical `executeAgentTaskInPersistentSession` copy-chain first run + persisted retry;
- sheet.move primitive and standalone agent-grade move_sheet, including retry;
- clear_range standalone agent-grade task, including retry;
- range.copy primitive and agent-grade copy_range, including coordinate-bound v3 persisted retry;
- range.move primitive and hardened agent-grade move_range, including source+target OCC retry and overlap rejection;
- insert_rows, delete_rows, insert_columns, delete_columns under the documented usedRange full-span structural contract, including values, formulas and coordinate-bound persisted retry;
- format_range standalone agent-grade task for the measured semantic contract (bold, italic, font name/size, horizontal alignment, wrap and number format), including persistent-session save barrier and proven persisted retry no-op. Formatting readback uses the live editor model where the public ApiRange exposes setters but not corresponding getters.

## Fail-closed hardening already covered

- rename graph source/target conflicts and cycles;
- overlapping generic writes after rename canonicalization;
- copy dependency gates;
- delete dependency gates;
- move_sheet membership-sensitive verification;
- range.move identical/overlapping source-target rejection;
- range.move target pre-state OCC conflict detection;
- structural coordinate-bound fingerprints;
- malformed/jagged/coordinate-less structural observations rejected;
- range.copy coordinate-bound v3 fingerprints;
- copy-chain expanded usedRange gaps must be semantically blank;
- format_range refuses properties without a semantic verification contract (including border in the current agent-grade task contract) and fails closed if requested formatting cannot be measured.

## Current contract limitations, not hidden PASS

### Sheet copy

The current sheet-copy verifier proves unique source/target identity, usedRange extent and cell content/formulas over that range. It does **not** prove complete worksheet equivalence for formatting, row/column dimensions, drawings/charts, validation, names or every other worksheet object. Therefore `copy_sheet` TRUE LIVE PASS is explicitly limited to the current narrow semantic contract.

### Structural operations

The structural four-operation contract is the measured full-span operation over current usedRange. It must not be described as proof of unrestricted entire-worksheet row/column semantics.

### Formatting

The current `format_range` TRUE LIVE PASS covers only properties for which the deployed editor provides stable semantic readback. It is not a blanket M4 formatting/layout/conditional-formatting PASS. Border formatting is deliberately rejected by the current agent task contract until semantic readback is proven; conditional formatting and broader layout operations remain separate migration work.

## Power User capabilities not yet migrated into the agent task vocabulary

The repository already has LIVE implementations/acceptances for capabilities outside the current W0.12-style agent task vocabulary. They are not automatically agent-grade merely because their earlier Power User acceptance is green. Remaining migration candidates include:

- remaining M4 layout/conditional-formatting families beyond the accepted `format_range` subset;
- M5 chart families;
- M6.1 sort/filter;
- M6.2 data validation;
- M6.3 defined names;
- M6.4 pivot tables;
- M7 protected-range ACL core;
- freeze panes and other earlier layout capabilities.

Each candidate needs an explicit task intent contract, fresh identity/read semantics, semantic verifier, whole-task proof, persistence behavior and retry/no-op/conflict semantics before being marked agent-grade.

## Runtime-deferred / unsupported capabilities

These remain governed by `EXCEL-POWER-USER-CAPABILITY-GAPS.md` and must not be promoted by the migration:

- M4.5 genuine Excel Table/ListObject identity on deployed 9.3.4.60: DEFERRED;
- M5.4 chart copy/duplicate: runtime UNSUPPORTED;
- M7 classic whole-sheet password protection: runtime UNSUPPORTED;
- M7 protected-range SetRange and SetAnyoneType semantic readback: DEFERRED;
- M7 protected-range delete: runtime UNSUPPORTED.

## Next migration order

1. Remove/deprecate the now-redundant dedicated copy-chain persistent wrapper after compatibility/static checks.
2. Continue M4 with bounded layout capabilities whose state can be semantically read back; keep conditional formatting separate.
3. Migrate sort/filter, data validation and defined names one family at a time into explicit agent-grade contracts.
4. Keep charts, pivots and protected-range ACL as explicit object-identity tasks with their own verifiers rather than reducing them to callback success.
5. Finish with a cross-capability persistent-session acceptance and update the capability gap register without converting documented runtime blockers to implementation failures.
