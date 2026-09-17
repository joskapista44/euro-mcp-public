# XLSX agent-grade persistent execution migration status

Audit date: 2026-09-17
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
- format_range standalone agent-grade task for the measured semantic contract (bold, italic, font name/size, horizontal alignment, wrap and number format), including persistent-session save barrier and proven persisted retry no-op. Formatting readback uses the live editor model where the public ApiRange exposes setters but not corresponding getters;
- M4.3 row/column layout: column width, row height, hide/show columns and rows, each with same-session semantic readback, persistence barrier and persisted no-op retry;
- M4.3 AutoFit columns and rows using the measured direct-range `AutoFit(bRows,bCols)` runtime contract, with operation-bound before/post dimensions, persistence barrier and persisted retry no-op;
- M4.3 insert/delete rows and columns through the already accepted agent-grade structural operations, under the documented usedRange full-span contract.
- M4 integrated agent-grade persistent execution on FILE_ID 1231187 at tested HEAD `0d06fc9c63bff528d875b4bb1554d7e3b148499f`: extended formatting, border, layout, merge, conditional formatting and freeze all passed primitive `LIVE_VERIFY`, followed by a fresh same-session whole-task verification where all six effects classified as satisfied. Task invocation 1 used exactly one editor session, performed 8 verified writes, used one successful persistence barrier and closed. The separate persisted retry invocation used exactly one new editor session, classified all six effects as no-op, performed 0 writes, used no persistence barrier and closed. Acceptance: `XLSX M4 INTEGRATED PERSISTENT LIVE ACCEPTANCE: PASS`, exit 0.
- `unmerge_range` dedicated TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `3867e0c77d882edde674460b90ee126a6d73d93d`: task invocation 1 used one editor session, performed fixture preparation plus merge/unmerge as 4 verified writes, completed fresh same-session whole-task verification, used one successful persistence barrier and closed. The separate idempotent retry used one new editor session, performed 0 writes, completed whole-task verification, used no barrier and closed. Acceptance: `XLSX PERSISTENT UNMERGE LIVE ACCEPTANCE: PASS`.

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
- format_range refuses properties without a semantic verification contract (including border in the current agent-grade task contract) and fails closed if requested formatting cannot be measured;
- AutoFit retry is bound to the operation plus measured before/post dimensions and fails closed if the current dimension matches neither accepted state.

## Integrated M4 TRUE LIVE acceptance

The previously pending extended-format, border, merge/unmerge, conditional-format and freeze contracts are now TRUE LIVE accepted for the exact tested semantic surface. The integrated result proves the W0.12 invariant per task invocation: one task = one persistent editor session, with same-session readback and whole-task verification. The persisted retry is a second task invocation, not a second session inside the first task.

## Current contract limitations, not hidden PASS

### Sheet copy

The current sheet-copy verifier proves unique source/target identity, usedRange extent and cell content/formulas over that range. It does **not** prove complete worksheet equivalence for formatting, row/column dimensions, drawings/charts, validation, names or every other worksheet object. Therefore `copy_sheet` TRUE LIVE PASS is explicitly limited to the current narrow semantic contract.

### Structural operations

The structural four-operation contract is the measured full-span operation over current usedRange. It must not be described as proof of unrestricted entire-worksheet row/column semantics.

### Formatting

The current `format_range` TRUE LIVE PASS covers only properties for which the deployed editor provides stable semantic readback. Extended color/vertical-alignment, borders, `merge_range`, the tested M4.6 conditional-format rule surface and freeze-at-range are TRUE LIVE accepted by the integrated run. This is not a blanket acceptance of every possible formatting, conditional-formatting, freeze, or merge variant. `unmerge_range` is also TRUE LIVE accepted by its dedicated persistent-session acceptance.

## Power User capabilities not yet migrated into the agent task vocabulary

The repository already has LIVE implementations/acceptances for capabilities outside the current W0.12-style agent task vocabulary. They are not automatically agent-grade merely because their earlier Power User acceptance is green. Remaining migration candidates include:

- the remaining M4 formatting/structure/conditional-formatting/freeze-panes work listed above;
- M5 chart families;
- M6.1 sort/filter;
- M6.2 data validation;
- M6.3 defined names;
- M6.4 pivot tables;
- M7 protected-range ACL core.

Each candidate needs an explicit task intent contract, fresh identity/read semantics, semantic verifier, whole-task proof, persistence behavior and retry/no-op/conflict semantics before being marked agent-grade.

## Runtime-deferred / unsupported capabilities

These remain governed by `EXCEL-POWER-USER-CAPABILITY-GAPS.md` and must not be promoted by the migration:

- M4.5 genuine Excel Table/ListObject identity on deployed 9.3.4.60: DEFERRED;
- M5.4 chart copy/duplicate: runtime UNSUPPORTED;
- M7 classic whole-sheet password protection: runtime UNSUPPORTED;
- M7 protected-range SetRange and SetAnyoneType semantic readback: DEFERRED;
- M7 protected-range delete: runtime UNSUPPORTED.

## Next migration order

1. Keep M4.5 explicitly runtime-DEFERRED unless the deployed runtime changes; the current implementable M4 agent-grade scope is TRUE LIVE accepted.
2. Migrate sort/filter, data validation and defined names one family at a time into explicit agent-grade contracts.
3. Keep charts, pivots and protected-range ACL as explicit object-identity tasks with their own verifiers rather than reducing them to callback success.
4. Finish with a cross-capability persistent-session acceptance and update the capability gap register without converting documented runtime blockers to implementation failures.
