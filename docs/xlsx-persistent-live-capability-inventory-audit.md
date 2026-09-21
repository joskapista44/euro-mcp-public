# XLSX persistent-LIVE capability inventory audit

Branch: `feature/xlsx-visual-demo`

## Audit standard

A capability is **agent-grade composed** only when the current batch path can execute it inside the owning persistent editor session and the batch can perform read-only whole-task verification. A standalone LIVE acceptance is evidence for the primitive, but is not by itself evidence that the capability is composable through `office_xlsx_batch`.

Required properties:

1. persistent editor session
2. LIVE postcondition verification
3. owning-session persistence barrier after writes
4. read-only whole-task verification
5. identical retry can become zero-write no-op
6. public batch/MCP exposure where the capability is intended for general agent use

## Composed families already present

The current `xlsx-persistent-batch.cjs` composes:

- core sheet create / write / copy / rename / delete
- clear range
- sheet move
- format
- fixed row/column layout
- merge / unmerge
- freeze / unfreeze
- sort
- filter / clear filter
- validation / clear validation
- defined names create / rename / delete
- conditional formatting add / delete
- pivot create / refresh / delete
- chart set / rename / delete
- page layout
- print setup
- print titles
- print area
- header/footer
- manual page breaks
- first-page number

The public `office_xlsx_batch` schema exposes those same general families. Page/Print has additionally passed the integrated one-task persistent LIVE closure.

## Capabilities with standalone persistent-LIVE evidence but not composed in the current batch

### A. AutoFit

Evidence exists in `test_xlsx_persistent_layout_autofit_live_acceptance.cjs`, but the batch planner explicitly rejects AutoFit with `xlsx-batch-autofit-not-supported` because operation-bound retry tokens have not been composed.

**Audit result: GAP — primitive evidence exists; agent-grade batch composition missing.**

### B. Structural row/column operations

Standalone structural tests exist for insert/delete rows and columns, formula preservation, retry behavior, full-span/runtime semantics and verifier behavior. The current batch family table does not expose structural intents.

Relevant evidence includes:

- `test_xlsx_persistent_structural_live_acceptance.cjs`
- `test_xlsx_persistent_structural_retry_live_acceptance.cjs`
- `test_xlsx_persistent_structural_formula_acceptance.cjs`
- `test_xlsx_structural_execution_static.cjs`

**Audit result: GAP — strong primitive/retry evidence exists; batch/MCP composition missing.**

### C. Range copy

Standalone persistent and agent-task evidence exists, including retry and verifier tests, but range-copy is not a family in the current batch planner/MCP schema.

Relevant evidence includes:

- `test_xlsx_persistent_range_copy_task_live_acceptance.cjs`
- `test_xlsx_range_copy_retry_formula_regression.cjs`
- `test_xlsx_range_copy_verifier_static.cjs`

**Audit result: GAP — primitive/retry evidence exists; batch/MCP composition missing.**

### D. Range move

Standalone persistent/retry/verifier evidence exists, but range-move is not a family in the current batch planner/MCP schema.

Relevant evidence includes:

- `test_xlsx_persistent_range_move_live_acceptance.cjs`
- `test_xlsx_persistent_range_move_retry_live_acceptance.cjs`
- `test_xlsx_range_move_verifier_static.cjs`

**Audit result: GAP — primitive/retry evidence exists; batch/MCP composition missing.**

## Deliberately unsupported

### Page order

The runtime exposes serialization/read state, but the investigation did not establish a verified persistent public write path. It remains intentionally unsupported rather than being exposed with weaker guarantees.

## Audit conclusion

The current workbook agent surface is already broad. The audit found **four composition gaps**, not a need for a new architecture:

1. AutoFit
2. structural row/column operations
3. range copy
4. range move

These are the only families identified in this audit that already have meaningful standalone implementation/evidence but are absent from the current agent-grade batch composition. Work should therefore proceed by composing and accepting these four existing capabilities, without redesigning already-closed families.

After these four gaps close, proceed to the cross-family complex agent-task acceptance, then the visual showcase, then final regression/closure and merge preparation.
