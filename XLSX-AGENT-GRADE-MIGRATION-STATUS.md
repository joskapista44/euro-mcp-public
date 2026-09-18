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

Migration proceeds one semantic family at a time. Sort, exact single-field `filter_range`, `clear_filter`, the durable M6.2 validation core and M6.3 defined names are now agent-grade TRUE LIVE accepted. The old direct-Playwright acceptances must not be used as agent-grade authority. M6.4 pivot create/delete and persisted retries have now passed the narrow acceptance below.

## Migration closure and runtime-bound exclusions

The implementable Excel Power User scope is migrated and TRUE LIVE accepted under the W0.12 execution model. Semantic pivot refresh is now part of the MCP vocabulary. No additional capability may be promoted without an exact public identity/readback contract.

Broader M6.4 pivot source/field identity remains runtime-bound until new public identity getters become available. M7 protected-range ACL remains primitive-only and runtime-DEFERRED for W0.12 because the deployed public API has no proven protected-range address getter. Chart set/rename/delete, including measurable presentation, position and series state, are W0.12 TRUE LIVE accepted; chart copy is runtime-UNSUPPORTED.

Operation-bound retry-token families (AutoFit, range move/copy and structural insert/delete) remain accepted standalone tasks rather than natural-retry batch intents. Exposing them through MCP later requires an explicit retry-token round trip; silently replaying them is forbidden.

## Runtime-deferred / unsupported capabilities

These remain governed by `EXCEL-POWER-USER-CAPABILITY-GAPS.md` and must not be promoted by the migration:

- M4.5 genuine Excel Table/ListObject identity on deployed 9.3.4.60: DEFERRED;
- M5.4 chart copy/duplicate: runtime UNSUPPORTED;
- M7 classic whole-sheet password protection: runtime UNSUPPORTED;
- M7 protected-range W0.12 identity: DEFERRED. Named-user ACL getters are measurable, but the deployed public API has no proven protected-range address getter. Therefore create/retry cannot prove that an existing titled object protects the requested cells; title plus ACL identity is insufficient for agent-grade no-op classification;
- M7 protected-range SetRange and SetAnyoneType semantic readback: DEFERRED;
- M7 protected-range delete: runtime UNSUPPORTED.

## Closure and future recheck policy

1. Keep the current implementable scope closed as TRUE LIVE accepted.
2. Recheck M4.5, broader pivot identity and M7 only when the deployed DocumentServer public API changes.
3. Keep charts and pivots as explicit object-identity tasks with their own verifiers rather than reducing them to callback success.
4. Preserve explicit retry-token semantics if destructive/relative operations are later exposed through MCP.

## Pivot runtime result and limits

User-reported acceptance at `eabb8bc0b472ce60b690a753ae4014e1bdf85993`: `XLSX PERSISTENT PIVOT LIVE ACCEPTANCE: PASS`. Create used 3 writes including fixture setup; delete used 1 write; both persisted retries used 0 writes. Each invocation used one editor session, with a save barrier only after writes. The tested pivot had one row, column and data field, PivotStyleMedium2, and GetData assertions 10/30/20. Source worksheet names are restricted to API-safe identifiers.

This proves the tested create/delete fixture, not unrestricted pivot identity: current verification measures source coordinates, field counts and sampled values, not complete source-sheet/field identity. Delete removes the identified parent worksheet; it does not prove that the sheet contains no unrelated content.

## Shared persistent batch executor — STATIC PASS and TRUE LIVE PASS

`xlsx-persistent-batch.cjs` composes core sheet/write operations, format, layout (excluding AutoFit), merge/unmerge, freeze, sort, filter, durable validation, defined names, conditional formatting, pivots and named charts in one existing editor session. It validates every operation before dispatch, conservatively rejects multiple goals within the same family/sheet or workbook name, and performs final verification through adapters that block mutation calls. The owning persistent-session wrapper alone saves and closes. The batch is not transactional: earlier writes may already exist if a later step fails.

Static acceptance covers preflight rejection, conflicting goals, zero-write retry and blocked writes during final verification. Successive LIVE gates cover the original seven-family batch, the nine-goal core/conditional-format/pivot expansion and the chart family. Existing family verification limitations still apply.

## MCP integration — protocol STATIC PASS and initial TRUE LIVE PASS

The packaged `euro-mcp-m44.cjs` entrypoint now registers `office_xlsx_batch` through `xlsx-batch-mcp.cjs`. Input is `file_id` plus `operations`; strict field validation and complete batch planning precede caller/credential resolution. Existing `coedit.detectCallerId` allowlisting and `credentialsFor` supply the identity and credentials. The tool dispatches once to the accepted persistent batch wrapper. Failed outcomes set MCP `isError`; thrown exceptions are not serialized because they may contain authenticated URLs. Failure is not rollback.

`test_xlsx_batch_mcp_static.cjs` passed using actual MCP client/server in-memory transport, covering tools/list, input rejection, preflight rejection before credential access, denied identity, missing credentials, single dispatch and safe error responses. The packaged entrypoint loads successfully. The initial stdio runtime gate passed through the real MCP entrypoint: a unique workbook name was created in one editor session with 1 write and one successful save barrier; the second MCP call reopened the workbook, proved the name already satisfied, used 0 writes and no barrier. It uses the existing elliot vault credential only inside the acceptance runner, passing it to the child's supported credential environment; production code has no vault dependency. The fixture leaves its uniquely named reference in the acceptance workbook. Measured open/task times were 3346/1113 ms for apply and 2860/66 ms for retry.

Example tool arguments: `{"file_id":"1231187","operations":[{"intent":"format_range","sheet":"Sheet1","range":"A1:B2","format":{"bold":true}}]}`. Production configuration uses `EURO_AGENT_ID`, the existing `EURO_COEDIT_AGENTS` allowlist, `EURO_COEDIT_NC_URL` and the caller-specific Nextcloud credentials. Do not use legacy primitive acceptances as evidence that every Power User capability has been migrated.

### Core editing expansion — STATIC PASS and TRUE LIVE PASS

The same `office_xlsx_batch` tool now also accepts `create_sheet`, `write_range`, `copy_sheet`, `rename_sheet` and `delete_sheet`. Core operations are planned together by the previously accepted generic agent task and must form a prefix before enhanced operations; this preserves create/write/rename dependency semantics while still using one editor session and one final read-only batch verification. MCP schemas reject intent-inapplicable fields. Static acceptance covers create+write, persisted no-op classification, ordering rejection and real MCP dispatch.

The expanded stdio acceptance passed: create sheet → write range → rename sheet → format header → set defined name ran in one MCP task with exactly 5 verified writes and one successful save barrier. The second MCP invocation reopened the workbook and proved all five goals with 0 writes and no barrier. Measured open/task times were 3080/1580 ms for apply and 2761/302 ms for retry. Copy/delete remain covered by their earlier dedicated TRUE LIVE acceptances but are not re-exercised in this combined fixture.

### Wider MCP batch — STATIC PASS and TRUE LIVE PASS

The MCP vocabulary now additionally composes `move_sheet`, `clear_range`, defined-name rename, conditional-format add/delete and the narrow accepted pivot create/delete contract. Original task indexes are preserved in batch-level verification reports (fixing the prior cosmetic 0 indexes on enhanced operations). Static tests cover the conditional-format adapter, write blocking during final verification and original-index mapping.

The broad runtime gate passed nine goals in one MCP request: create, write, rename, move, clear, format, conditional format, defined name and pivot. Final-state canonicalization removed the dead intermediate value covered by the later clear. The first invocation used one editor session, 8 verified writes, read-only whole-task verification and one successful save barrier. The persisted retry used one editor session, proved all nine goals, performed 0 writes and used no barrier. Measured open/task times were 3350/2040 ms for apply and 2843/696 ms for retry. Range move and structural insert/delete are intentionally not exposed through this natural-retry batch yet because their accepted contracts require operation-bound retry tokens; replaying them without the returned token could repeat a destructive mutation. AutoFit has the same token-bound constraint. Named chart operations were added and accepted in the subsequent chart gates.

### Wider MCP batch first runtime finding

The first nine-goal run executed all nine mutations successfully but correctly failed closed before the persistence barrier during batch-wide final verification. Cause: the core verifier still expected the original `write_range` value in a cell intentionally overwritten later by `clear_range`. This was an orchestration final-state bug, not an ONLYOFFICE mutation failure. The batch planner now derives a separate read-only core verification projection in which downstream clear intersections are blank, while the apply phase still uses the original requested write. A static regression asserts the apply/verify values differ exactly at the overlaid cell (`4` then `null`). The later acceptance below closes this finding. Because the failed run had writes but no successful persistence barrier, it is not recorded as acceptance PASS.

### Wider MCP batch second runtime finding

After the read-only projection fix, the first invocation passed all nine final checks, used one editor session, performed 9 writes and completed the save barrier. The persisted retry then failed closed with 0 writes because its apply-phase core classifier still compared against the overwritten intermediate value. This confirmed that projection only during final verification was insufficient.

The final-state projection is now used for both dispatch planning and final verification. A core value later covered by `clear_range` is a dead intermediate write and is canonicalized to blank before dispatch. The expected first-run write count is therefore 8, not 9; retry remains 0. Static regression now requires `[null,null]` for apply and verify projections. The later nine-goal PASS closes the retry gate; the earlier incomplete invocation is not treated as acceptance evidence.

### Nine-goal MCP acceptance result

User-reported result: `XLSX BATCH MCP LIVE ACCEPTANCE: PASS`. All batch-level checks preserve the original operation indexes 0–8. This closes the shared MCP orchestration gate for that nine-goal vocabulary. Natural-retry exposure of operation-bound token families remains intentionally excluded; documented runtime gaps remain exclusions, not failures.

### Chart MCP core acceptance result

User-reported result at `0d65732`: `XLSX CHART MCP LIVE ACCEPTANCE: PASS`. Named chart creation plus fixture setup ran in one editor session with 3 writes, read-only whole-task verification and one successful save barrier; persisted retry reopened the workbook, performed 0 writes and used no barrier. Named deletion ran in one session with 1 write and a successful barrier; its persisted retry used 0 writes and no barrier. The editor-key deletion fallback now uses bounded semantic readback polling rather than a fixed post-key delay. This proves the basic name/type/title/size/series-count contract.

User-reported advanced result at `2110aac`: `XLSX CHART MCP LIVE ACCEPTANCE: PASS`. The same `set_chart` task additionally proved persisted legend position, horizontal/vertical axis titles, data-label flags, chart style, exact drawing position, series name, values formula and category formula. Apply used one session with 3 task-level writes and a successful barrier; reopen retry used 0 writes and no barrier. Delete and delete retry remained 1/0 writes. Measured open/task times were 3106/1599 ms (apply), 2946/245 ms (retry), 2906/1139 ms (delete) and 2777/53 ms (delete retry). The initial advanced attempt exposed that `ApplyChartStyle` resets presentation state; mutation ordering is now style → series/position → legend/axes/labels, with static regression coverage.

User-reported chart lifecycle result at `55dcd4c`: `XLSX CHART MCP LIVE ACCEPTANCE: PASS`. Exact named rename used one editor session, 1 verified write and one successful save barrier; persisted rename retry used 0 writes and no barrier. Measured open/task times were 2858/1142 ms for rename and 2843/68 ms for retry. Together with the set/delete gates this closes every implementable M5.4 chart operation; copy/duplicate remains documented runtime-UNSUPPORTED.

### Pivot refresh MCP — STATIC PASS and TRUE LIVE PASS

`refresh_pivot` uses the existing unique pivot name plus source-coordinate/style/field-count fingerprint and required public `GetData` assertions. If those assertions already match, retry is a proven no-op. Otherwise the observer rechecks unchanged live state, calls public `RefreshTable`, verifies the requested semantic totals, runs the batch-wide read-only verification, then lets the single owning session perform one save barrier. No fixed delay or second editor session is introduced.

User-reported result at `21ecbc8`: `XLSX PIVOT REFRESH MCP LIVE ACCEPTANCE: PASS`. Setup used one session and 3 writes. The refresh invocation changed one source value and refreshed the pivot in one session with exactly 2 writes and one successful save barrier; persisted retry reopened once, proved both source value and pivot totals already satisfied, used 0 writes and no barrier. Cleanup removed the generated pivot sheet with 1 write. Measured open/task times were 3086/1371 ms (setup), 2744/1220 ms (refresh), 2738/248 ms (retry) and 2681/1191 ms (cleanup).

### Final Excel Power User MCP gate — TRUE LIVE PASS

The final gate composes 13 natural-retry-safe goals in one real `office_xlsx_batch` call: sheet move, range clear, formatting, fixed layout, merge, freeze, sort, filter, durable validation, defined name, conditional formatting, pivot creation and an advanced named chart. It then reopens the same final-state request and requires 0 writes and no save barrier. Operation-bound retry-token families remain deliberately separate; runtime-deferred/unsupported capabilities remain exclusions rather than false failures.

The first final-gate run completed all 13 mutations but correctly failed closed before the save barrier during read-only whole-task verification. The failing goal was freeze panes: its immediate mutation readback had passed, but later pivot/chart object work changed the editor sheet/view context, so final freeze readback no longer matched and the read-only adapter blocked a rewrite. The batch planner now schedules freeze goals last for both apply and final verification while preserving their original operation indexes. A static regression covers this ordering rule; no delay or extra editor session was added.

The reordered final gate passed its 13-goal apply and persisted 0-write retry. Its separate object cleanup then exposed another sheet-context dependency: deleting the pivot worksheet first changed the active sheet, while the chart fallback selected the named chart through its worksheet object but sent the editor Delete key to the wrong active context. The fallback now uses public `ApiWorksheet.SetActive`, verifies the active sheet name when readable, then performs public Select, Delete-key dispatch and bounded public inventory readback.

User-reported final result at `c2f56b5`: `XLSX FINAL POWER USER MCP LIVE ACCEPTANCE: PASS`. Fixture setup used one editor session, 4 writes and one save barrier. The 13-goal Power User request used one editor session, exactly 13 verified task-level writes, read-only whole-task verification and one successful save barrier. The persisted retry reopened the workbook once, proved every final state, used 0 writes and no barrier. Pivot-plus-chart cleanup used one editor session, 2 writes, read-only whole-task verification and one save barrier. Original operation indexes were preserved even though freeze was safely scheduled last. Measured open/task times were 3251/1631 ms (setup), 2859/2719 ms (apply), 3127/851 ms (retry) and 3205/1349 ms (cleanup).

This closes the current Excel W0.12 MCP migration. Every supported combined task follows one editor session, fresh live reads around mutation, semantic readback, read-only whole-task verification, a single save barrier only when writes occurred, and close. No arbitrary delay or second editor session is used inside a task.

### Agent-driven visual showcase — STATIC PASS, TRUE LIVE pending

`XLSX-VISUAL-SHOWCASE-AGENT-TASK.md` is a self-contained natural-language assignment for an agent using the MCP, not a hard-coded execution script. After the operator opens the workbook and replies `INDULHAT`, the agent must compose one `office_xlsx_batch` call that creates a polished dashboard, formula-linked monthly plan, filtered/validated transaction sheet, conditional formatting, a semantic pivot and two named native charts. The result is deliberately left in the workbook for visual review.

The batch planner now distinguishes independent format, fixed-layout, merge and named-chart targets on the same worksheet, so a rich document remains one task and one editor session. Exact duplicate targets still fail closed. Optional bounded `readbacks` return selected final ranges after read-only whole-task verification and before the owning session performs its single save barrier. This lets the agent report live calculated KPI values from the same session instead of inferring success from mutation callbacks. Static acceptance covers the 52-operation showcase contract, 51 expected fresh-state writes, three same-session readbacks, duplicate-target rejection and MCP schema/transport forwarding. TRUE LIVE acceptance remains pending.

Open-editor diagnosis found that `Api.AddSheet()` delegates to the collaboration-aware `asc_addWorksheet()` lock path asynchronously. The old adapter verified sheet identity inside the same `callCommand` callback, which raced the lock callback whenever another editor was already present. Native create and move dispatch now wait on live worksheet inventory state in the owning session, without a second browser session or fixed delay. Static acceptance simulates delayed collaboration-lock completion; `test_xlsx_open_coedit_create_live_acceptance.cjs` is the dedicated TRUE LIVE gate while the operator keeps the target workbook open.

User-reported TRUE LIVE result at `2102a32` on `mcp_test.xlsx` (`file_id=1236770`): `XLSX OPEN COEDIT CREATE/MOVE LIVE ACCEPTANCE: PASS`. With the human editor kept open, one agent editor session created `EURO_COEDIT_1074444`, wrote and live-read `A1:B2`, then moved the sheet immediately before `Sheet1`. All three whole-task checks passed, the same-session readback carried `authority=LIVE_READ`, exactly 3 writes were counted, and the single force-save barrier completed in 0 ms before the wrapper closed the session. This disproves the earlier Nextcloud-file-lock hypothesis for this failure and confirms the asynchronous collaboration-lock race as fixed.

The first agent-driven showcase attempt (`run_id=4COBEG`) then failed closed after exactly one write, before any persistence barrier or final readback. The request interleaved `create_sheet` and `write_range`; dashboard cell `B4` referenced the not-yet-created plan worksheet, so formula identity readback returned an empty formula and correctly reported `formula-write-not-verified`. This was a deterministic orchestration-order defect, not a transient editor failure.

The batch planner now performs a stable creation-first normalization inside the core prefix: every `create_sheet` executes before any `write_range`, while original request indexes remain attached to receipts and verification. A regression reproduces the dashboard/plan/data interleaving and requires execution order `create Dash, create Plan, create Data, write Dash, write Plan, write Data`. The natural-language showcase contract also explicitly requires those three creations before all three writes. Targeted core, batch MCP and showcase tests plus the full repository suite pass. A fresh showcase TRUE LIVE run remains pending.

The next showcase attempt (`run_id=768SVZ`) passed the creation-first boundary and reached operation 44 after 43 verified writes in one session. It then failed closed at `set_defined_name`: `AddDefName` dispatched without error, but the immediate `GetDefName` read still reported the name absent. No final readbacks or persistence barrier ran, and later conditional-format, pivot, chart and freeze goals were not dispatched. This is the same collaboration-lock visibility class previously proven for worksheet creation. Defined-name creation now dispatches once and performs bounded, state-driven public `GetDefName` polling in the same editor session; it never retries the mutation. Static regression requires eventual semantic visibility and exactly one mutating dispatch. The idempotent 768SVZ retry is the pending TRUE LIVE gate.

The first 768SVZ retry then failed closed before reaching the defined name. The initial run had applied currency formatting after the core write; on retry the live getter returned the twelve unit-price numbers as numeric strings (`"190"`) with currency number format and displayed text (`$190`). The verifier compared JavaScript primitive types and rejected all twelve even though the numeric values were identical. Range verification and core no-op classification now share a guarded semantic numeric comparison: stringified getter output is accepted for a finite expected number only when an explicit non-General, non-text number format produces a formatted display distinct from the stored numeric string. A genuine General/text `"190"` remains different from numeric `190`. Static tests cover primitive verification and zero-write core retry. The same 768SVZ final-state retry remains the pending TRUE LIVE gate.
