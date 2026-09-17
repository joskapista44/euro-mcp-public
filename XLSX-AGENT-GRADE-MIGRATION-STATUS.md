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
- M6.1 `sort_range` agent-grade TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `851e55c4cb33c9770533bd2ab62902e53570a041`: task invocation 1 used one editor session, performed create/write/sort as 3 verified writes, proved exact ordered keys `Alpha, Bravo, Charlie, Delta`, completed same-session whole-task verification, used one persistence barrier and closed. The separate retry used one new editor session, classified the requested ordering as already satisfied, performed 0 writes, used no barrier and closed. Static regression and LIVE acceptance both PASS.
- M6.1 `filter_range` agent-grade TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `3f0c900e220662f2627fd2af93eb9f3af7ca00b5`: exact `A1:C5` range, field 3, `xlOr`, criteria `A` were proven from the live model; task invocation 1 used one editor session, performed create/write/filter as 3 verified writes, completed whole-task verification, used one persistence barrier and closed. The separate retry proved the identical filter already satisfied, performed 0 writes, used no barrier and closed. Static regression and LIVE acceptance both PASS.
- M6.1 `clear_filter` agent-grade TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `b215220b76bcc710ce68ae8c86cbf1576d73c6b5`: task invocation 1 created and populated the fixture, applied an exact filter, then cleared its active criteria through `ShowAllData` in one editor session. Exact range `A1:C5` remained present with zero active filters, 4 verified writes used one persistence barrier, and same-session whole-task verification passed. The separate retry proved the cleared state already satisfied, performed 0 writes, used no barrier and closed. `filterMode: true` is expected because ONLYOFFICE preserves the AutoFilter controls after clearing criteria. Static regression and LIVE acceptance both PASS.
- M6.2 `set_validation` and `clear_validation` agent-grade TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `f519941ed67cb438e69ce30cf0907de9ca8c07c1`: the durable validation core (type, alert style, operator and formulas) passed exact public-getter readback, one-session mutation, one save barrier and persisted zero-write retry. Clear and its persisted retry also passed. The deployed runtime does not persist non-default `ignoreBlank`, input/error title or message values; those options are therefore rejected by the planner rather than silently rewritten on every open.
- M6.3 `set_defined_name`, `rename_defined_name` and `delete_defined_name` agent-grade TRUE LIVE acceptance on FILE_ID 1231187 at tested HEAD `f519941ed67cb438e69ce30cf0907de9ca8c07c1`: each mutation used one persistent editor session, exact public `GetName`/`GetRefersTo` readback and one save barrier. Separate persisted retries for set, rename and delete each proved the final state with 0 writes and no barrier. Static regression and LIVE acceptance both PASS.

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

### Data validation

The W0.12 contract deliberately covers only the validation fields proven durable across save and reopen: type, alert style, operator and formulas. On the deployed runtime, non-default `ignoreBlank`, input/error titles and messages read back correctly in-session but reopen with defaults. They are explicit planner rejections, not hidden support.

## M6.1 sort/filter migration audit (2026-09-17)

The existing M6.1 implementation remains a valid primitive-level LIVE PASS, but it is not yet W0.12-style agent-grade execution:

- `m61-live-acceptance.cjs` performs its own browser login/editor navigation and uses fixed 2.5 s and 22 s waits instead of the minimal persistent editor-session wrapper;
- operations are called as pre-expanded low-level commands rather than through an explicit agent task planner;
- sort lacks fresh mutation-boundary range fingerprint comparison and persisted already-satisfied/no-op classification;
- filter enable readback does not require exact requested range identity;
- filter set readback proves an active readable filter but does not prove exact field/operator/criteria identity;
- reapply reports before/after state without a semantic requested-final-state assertion;
- there is no task-level whole-state verifier, change receipt, single save-completion barrier, or persisted retry contract.

Migration proceeds one semantic family at a time. Sort, exact single-field `filter_range`, `clear_filter`, the durable M6.2 validation core and M6.3 defined names are now agent-grade TRUE LIVE accepted. The old direct-Playwright acceptances must not be used as agent-grade authority. M6.4 pivots are next.

## Power User capabilities not yet migrated into the agent task vocabulary

The repository already has LIVE implementations/acceptances for capabilities outside the current W0.12-style agent task vocabulary. They are not automatically agent-grade merely because their earlier Power User acceptance is green. Remaining migration candidates include:

- the remaining M4 formatting/structure/conditional-formatting/freeze-panes work listed above;
- M5 chart families;
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
2. Migrate M6.4 pivot tables, then the implementable M7 protected-range ACL core.
3. Keep charts, pivots and protected-range ACL as explicit object-identity tasks with their own verifiers rather than reducing them to callback success.
4. Finish with a cross-capability persistent-session acceptance and update the capability gap register without converting documented runtime blockers to implementation failures.
