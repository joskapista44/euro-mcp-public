# EURO-MCP XLSX W0.12 visual showcase — session handoff

Date: 2026-09-18

## 1. Project goal

Convert the Excel capabilities of `joskapista44/euro-mcp-public` to the same agent-grade execution model used by the EURO-MCP Word W0.12 work.

The target is not merely an API that can mutate an XLSX file. The MCP must let an agent enter one live ONLYOFFICE co-editing session, observe the current workbook, apply a complete requested final state through ONLYOFFICE APIs, read the result back semantically, verify the whole task, persist only verified changes, and leave the session. A later identical invocation must prove that the persisted workbook already satisfies the request without rewriting it.

The user-facing proof is a polished sales-performance workbook built live in `mcp_test.xlsx` while the operator watches it change. The workbook demonstrates formulas, formatting, layout, sorting, filtering, validation, names, conditional formatting, pivot analysis, charts and freeze panes together. It must also demonstrate that the agent can read and report what it actually produced.

## 2. Repository and live environment

- Repository: `joskapista44/euro-mcp-public`
- Current development branch: `feature/xlsx-visual-demo`
- User worktree: `/home/user/xlsx-persistent-live-session-wt`
- Base implementation branch used earlier: `feature/xlsx-persistent-live-session`
- MCP entrypoint: `euro-mcp-m44.cjs`
- MCP tool: `office_xlsx_batch`
- Live workbook: `mcp_test.xlsx`
- Nextcloud file ID: `1236770`
- Showcase run ID: `768SVZ`
- Live caller: `elliot`
- Required environment: `EURO_AGENT_ID=elliot EURO_COEDIT_AGENTS=elliot`
- Nextcloud: `https://mt-server.eu`
- ONLYOFFICE/EuroOffice runtime: deployed 9.3.4 hotfix line with co-edit API

The credentials stay in the existing vault flow. Do not put passwords into commands, source files or reports.

## 3. User working rules

- Work directly on GitHub. Do not ask the user to perform source-control or editing work that can be done through the GitHub connector.
- Continue autonomously and stop only when a genuine live/runtime test must run in the user's environment.
- Give the user one short copyable block, normally `git fetch/merge` plus one Node command.
- Do not delegate this work to another agent.
- Do not use fixed sleeps or arbitrary delays.
- Do not introduce a second Chrome/editor session inside one task.
- Use the smallest possible Chrome/Playwright surface only to enter and own the editor session. Spreadsheet work and semantic observations must use the live ONLYOFFICE API/model.
- Do not weaken a verifier merely to turn a failing acceptance green.
- Do not claim completion from callback success, a screenshot, or a correct-looking workbook.

## 4. Mandatory W0.12 agent-grade criteria

Every supported task must satisfy all of the following.

### 4.1 One task, one editor session

- One MCP task opens exactly one authenticated live editor session.
- All reads, mutations, post-mutation observations and whole-task verification occur in that same session.
- The owning wrapper alone performs persistence and closes the session.
- A persisted retry is a separate task invocation with one new editor session, not a second session hidden inside the first task.

### 4.2 Minimal browser use and no fixed delay

- Browser automation is only the transport for authenticated editor entry and the live `callCommand` bridge.
- Spreadsheet actions use ONLYOFFICE APIs or measured live editor-model getters/setters.
- There are no fixed `sleep` delays. Waiting is allowed only for a concrete postcondition, readiness state or save-completion condition.

### 4.3 Plan before mutation

- Validate the complete request and reject unsupported, malformed, ambiguous or conflicting goals before dispatching writes.
- Establish target identity from fresh live inventory.
- Re-read the relevant state at the mutation boundary so a stale plan cannot overwrite a concurrently changed workbook.
- For cross-sheet formulas, create every referenced sheet before the first range write.

### 4.4 Authority and fail-closed behavior

The authority ladder is:

- `PLAN_ONLY`
- `DISPATCH_ONLY`
- `LIVE_READ`
- `LIVE_VERIFY`

Callback success or absence of an exception is never enough. If identity, state or a requested semantic property cannot be measured exactly, the operation must fail closed. Unknown state must not be converted into success or an assumed no-op.

### 4.5 Semantic readback

- After each mutation, read the requested semantic state back in the same live session.
- Verify values, formulas, ranges, names, object identity and supported presentation properties, not just object counts or callback results.
- Return a receipt only for a mutation that reached `LIVE_VERIFY`.
- A retry may classify `ALREADY_SATISFIED` only from live evidence of the exact requested final state.

### 4.6 Whole-task verification

- After all planned steps, run a fresh read-only whole-task verification.
- Verification adapters must be unable to dispatch writes.
- Every operation check must pass from the final state, including dependencies whose coordinates or worksheet names changed during the task.
- Requested final range readbacks must come from this same session and carry `LIVE_READ` authority.

### 4.7 Persistence

- Run the save/persistence barrier only after a successful task with at least one verified write.
- Do not save a failed or merely dispatched task as if it were accepted.
- A successful mutating invocation must prove the persistence barrier completed, then close the session.
- A zero-write invocation must not run a persistence barrier.

### 4.8 Persisted retry and idempotence

The identical request after save/reopen must have exactly one of two outcomes:

1. the entire requested final state is proven already satisfied, with `noOp=true`, `writes=0`, `persistenceBarrier=null`; or
2. the executor fails closed on a measured conflict or unverifiable state.

Repeated corrective writes on every reopen are not acceptable, even if every individual invocation ends with a correct-looking workbook.

### 4.9 Co-editing safety

- Use the authenticated caller identity and existing allowlist/credential mapping.
- Do not mutate unrelated pre-existing sheets or objects.
- Do not silently overwrite ambiguous same-name objects.
- The shared batch is not transactional. If a later operation fails, earlier verified writes may remain. Failures must report this honestly rather than claim rollback.

## 5. MCP batch contract

`office_xlsx_batch` receives one `file_id`, an ordered final-state operation list and optional final readbacks.

For the visual showcase:

- the agent makes one mutating MCP call containing the complete construction plan;
- all four `create_sheet` operations precede every `write_range` operation;
- core operations form a prefix before enhanced operations;
- the batch executor applies family tasks inside one already-open editor session;
- freeze operations are applied last because chart/pivot work may switch the active worksheet;
- final verification is read-only;
- the three requested range readbacks are returned in the same response.

The production agent must not use standalone write primitives, DocBuilder, downloaded XLSX manipulation or direct OOXML changes to complete this task.

## 6. Visual showcase being tested

The canonical specification is in:

- `XLSX-VISUAL-SHOWCASE-AGENT-TASK.md`
- `xlsx-visual-showcase-contract.cjs`
- `test_xlsx_visual_showcase_resume_live_acceptance.cjs`

Run-specific worksheets:

- `MCP_Dash_768SVZ`
- `MCP_Plan_768SVZ`
- `MCP_Data_768SVZ`
- `MCP_PivotView_768SVZ`

Run-specific objects:

- main defined name: `MCP_Sales_768SVZ`
- pivot-source identity: `MCP_PivotSource_768SVZ`
- pivot: `MCP_Pivot_768SVZ`
- charts: `MCP_Trend_768SVZ` and `MCP_Region_768SVZ`
- chart geometry identities:
  - `MCP_TrendSize_768SVZ_3800000x2300000`
  - `MCP_RegionSize_768SVZ_3800000x2300000`

The request currently contains 56 operations/checks and three final readbacks.

### 6.1 Data sheet

`A1:H13` contains twelve monthly/region/product records. Units, price and score are typed numbers. Revenue and status are formulas. The task also demonstrates header formatting, currency/integer formats, fixed column/row dimensions, sorting, active filtering, whole-number validation, conditional formatting, a defined source range and freeze panes.

### 6.2 Plan sheet

`A1:E9` links to live data-sheet formulas and calculates monthly actual revenue, target, variance and attainment. It demonstrates merged title layout, formatted sections, currency/percentage formats, positive-variance conditional formatting and freeze panes.

### 6.3 Dashboard

The dashboard has KPI formulas, regional summaries, two chart-source tables and two native editable charts. It is moved immediately before `Sheet1`.

Expected KPI values, which must still be reported from the live readback rather than trusted as constants:

- total revenue: `287205`
- target revenue: `282000`
- variance: `5205`
- attainment: `1.0184574468085106`

### 6.4 Pivot

The pivot has Region rows, Product columns and Units as the data field, style `PivotStyleMedium2`. Required semantic `GetData` assertions are:

- North/Core: 405
- North/Plus: 252
- South/Core: 300
- South/Plus: 200

### 6.5 Final evidence

A successful construction/resume response must prove:

- `ok=true` and `authority=LIVE_VERIFY`;
- one editor session and wrapper-owned close;
- a barrier only if verified writes occurred;
- 56 read-only final operation checks, all passing;
- three same-session final range readbacks;
- correct formulas and live KPI values;
- exact pivot assertions and chart identities.

The final W0.12 closure requires a subsequent strict invocation with:

- `noOp=true`;
- `writes=0`;
- `persistenceBarrier=null`;
- the same 56 passing checks and three readbacks.

## 7. What has already passed

The persistent XLSX work has TRUE LIVE coverage for the implemented semantic surfaces of:

- create/write/formulas/rename/delete/copy sheets;
- move sheet and clear range;
- range copy/move and structural row/column operations under their documented narrow contracts;
- formatting, extended colors/alignment, borders, fixed layout, AutoFit as standalone token-bound tasks, merge/unmerge, conditional formatting and freeze panes;
- sort, exact filter and filter clear;
- durable validation core and clear;
- defined-name create/update/rename/delete;
- pivot create, refresh and delete under the narrow measured identity contract;
- chart set, presentation, series, position, rename and delete under the measured contract;
- cross-capability persistent batch and MCP stdio integration;
- final power-user batch construction and persisted no-op retry for the earlier acceptance fixture.

The visual showcase itself has already proved the following on `mcp_test.xlsx`:

- the four sheets and their content exist;
- formulas calculate correctly;
- formatting, dimensions, sort, filter, validation, conditional formatting, names, pivot, charts and freeze goals pass final verification;
- all three final readbacks are present;
- the KPI values above are observed live;
- the batch completes with 56/56 final checks in one editor session.

Therefore the document construction and semantic readback goal is complete. The only open showcase gate is persisted zero-write idempotence.

## 8. Runtime findings encountered during the showcase

These findings are important because several failures were genuine new runtime boundaries rather than random test instability.

1. Sheet creation and formula order: dashboard formulas were initially written before their referenced sheets existed. The final builder creates all sheets first, then writes data, plan and dashboard.
2. Numeric readback: ONLYOFFICE getters return numeric cells as strings in this runtime. `WorksheetFunction.TYPE(range)` returns 16 for these cells, while independent `COUNT` and `SUM` reference controls prove numeric spreadsheet semantics. The range verifier now uses the measured reference semantics and still rejects numeric-looking text headers.
3. Defined names: validation and ordinary workbook names are durable through the explicit defined-name family. Silent or non-durable in-task shortcuts are not accepted.
4. Pivot source identity: reopened `GetSource()` can throw because the wrapper lacks worksheet state. The showcase uses the explicit durable pivot-source name and sampled `GetData` assertions.
5. Freeze panes: chart/pivot operations may change active-sheet context, so the batch orders freeze goals last.
6. Formula text: readback may contain harmless spaces such as `= SUM(...)`; formula verification canonicalizes formula identity instead of requiring byte-identical whitespace.
7. Chart geometry: after save/reopen, both named charts can report `GetWidth/GetHeight = 0x0` even though `SetSize` verified exact nonzero values in the mutating session. The chart state otherwise remains exact.

## 9. Current chart-geometry solution

Each chart plan derives a size-bearing geometry identity name and an exact absolute valid range reference. A durable identity is accepted only as fallback for the measured reopened `0x0` getter state. It never overrides conflicting nonzero public geometry, and a wrong existing marker remains a fail-closed conflict.

An earlier attempt created the marker internally during the chart task. A read-only probe proved that those markers were visible in-session but absent after reopen, so that path was rejected as non-durable.

The current contract uses two explicit `set_defined_name` operations immediately after the chart operations. This routes persistence through the already accepted defined-name family. The chart mutation may pass on exact nonzero public getters before the marker exists; batch-wide final verification then sees the explicit marker. A reopened `0x0` chart may pass only when its exact marker is present.

## 10. Exact current status

Latest meaningful live result before the diagnostic publication error:

```text
XLSX VISUAL SHOWCASE RESUME: PASS
runId=768SVZ
file_id=1236770
noOp=false
writes=1
oneEditorSession=true
barrier=true
checks=56
readbacks=3
KPI=287205 / 282000 / 5205 / 1.0184574468085106
```

This proves the workbook remains semantically complete, but it does not satisfy the final W0.12 retry criterion because one operation still wrote.

The immediately following strict invocation on commit `e3d933d` reached the same successful task result but rejected it because aggregate `noOp` was false. The test did not yet print the responsible step.

A diagnostic change was then published as commit `f659cca`, but the GitHub-published line was malformed:

```text
SyntaxError: Unexpected token 'null'
```

That syntax error occurs while Node parses the acceptance file, before connecting to the MCP or opening the workbook. Consequently this particular invocation performed no session, no read and no write. It says nothing new about workbook state.

The malformed publication is corrected in the same GitHub change that adds this handoff. The diagnostic filters `r.steps` for results whose `noOp !== true` and prints only their original index, intent, outcome, no-op flag, planned operation and verified final state.

## 11. Immediate continuation plan

1. Pull the handoff/diagnostic correction from `feature/xlsx-visual-demo`.
2. Ask the user to run only the strict retry, not another permissive completion run.
3. Read `XLSX VISUAL SHOWCASE ZERO-WRITE RETRY DIAGNOSTIC` and identify the one non-no-op family.
4. Fix only that measured family. Preserve exact fail-closed semantics and do not add sleeps or a second session.
5. Run the strict invocation again.
6. Close the showcase only when it prints `XLSX VISUAL SHOWCASE ZERO-WRITE RETRY: PASS` with `noOp=true`, `writes=0`, no barrier, 56 checks and three readbacks.
7. Record the final live evidence in `XLSX-AGENT-GRADE-MIGRATION-STATUS.md` and commit it.

## 12. Required next live command pattern

After the corrected diagnostic is on GitHub, give the user one copyable block only:

```bash
cd /home/user/xlsx-persistent-live-session-wt && git fetch euro-public feature/xlsx-visual-demo && git merge --ff-only euro-public/feature/xlsx-visual-demo
EURO_XLSX_REQUIRE_NOOP=1 EURO_AGENT_ID=elliot EURO_COEDIT_AGENTS=elliot node test_xlsx_visual_showcase_resume_live_acceptance.cjs 1>&2
```

If it fails, use the printed non-no-op step as evidence. Do not ask the user to run broad probes or repeat a known write without a code change.

## 13. Explicit limitations and open backlog

Do not describe these as completed capabilities:

- Print/page layout is still open. The MCP does not yet guarantee print area, orientation, paper size, margins, centering, repeated title rows/columns or fit-to-one-page output. This must become a separate agent-grade family with live readback and zero-write retry.
- Genuine Excel Table/ListObject identity is deferred on the deployed runtime.
- Chart copy/duplicate is unsupported.
- Whole-sheet password protection is unsupported.
- Protected-range address identity and deletion remain deferred/unsupported.
- Pivot acceptance proves the narrow source identity, field counts and sampled semantic values, not unrestricted pivot-table equivalence.
- Sheet copy and structural operations retain their documented narrow semantic contracts.
- Data-validation support deliberately excludes options that this runtime does not persist after reopen.
- AutoFit and relative/destructive token-bound operations must not be silently exposed as ordinary natural-retry batch operations without explicit retry-token semantics.

## 14. Completion definition

The visual showcase milestone is complete only when all of the following are true simultaneously:

- the live workbook is visually complete and semantically correct;
- the agent receives and reports same-session live readbacks;
- all 56 whole-task checks pass read-only;
- the strict persisted retry returns `noOp=true`;
- the strict persisted retry performs exactly zero writes;
- the strict persisted retry runs no persistence barrier;
- the session is closed by the wrapper;
- the final evidence is committed to GitHub.

Until then, report the state as: **visual showcase construction PASS; W0.12 persisted zero-write closure PENDING**.
