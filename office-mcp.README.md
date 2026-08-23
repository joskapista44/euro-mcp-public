# Office MCP — scaffold + design notes

Scoped MCP server for driving the euro-office / OnlyOffice Document Server (DS).

Status: `office_status`, **`xlsx_set_cells` and `xlsx_recalc` are implemented and tested**.
`xlsx_recalc` recalculates through **headless LibreOffice** (see below), not the Document
Server. **Not wired into any agent session yet** (that is a separate, access-gated step).

The workbook rewrite lives in `office_xlsx.py`, invoked from the Node server. xlsx is a ZIP
container, Node has no bundled ZIP support and this host has no `zip` binary — so the
alternative was hand-rolling central-directory writing, which is a reliable way to produce a
subtly corrupt workbook. `zipfile` is stdlib, so this adds no dependency.

## The design decision that shapes everything

Editing and recalculating are **two different capabilities with different dependencies**:

| | needs the Document Server? | blocked today? |
|---|---|---|
| write a cell value/formula | **no** — local OOXML rewrite | no |
| recalculate formulas | **no** — headless LibreOffice (was: DS docbuilder) | no |

So they are separate tools, and `office_status` exists to tell a caller which of the two is
actually available *before* it starts editing.

**ConvertService is not a recalc path.** It is a converter: for a formula cell it returns the
*cached* value, or the formula text when there is no cache. Routing recalc through it would
hand back stale numbers that look correct — the worst failure mode for a spreadsheet tool.
`xlsx_recalc` therefore refuses rather than falling back to it.

**How recalc works now.** Recalculation was pivoted off the Document Server (its docbuilder
API is `advanced_api`-licence-gated and was not available) onto **headless LibreOffice**. The
critical fact: LibreOffice does not recalculate xlsx on load by default and `--convert-to`
does not override it — verified on three versions (7.3.7.2, 7.4.7.2 and 26.2.4.2): a plain conversion
returns the stale cache (999) while a per-job profile with recalc-on-load forced returns the
computed value (5). `office_recalc.py` builds that profile per job; `office_recalc_gate.py` is
the standing regression that proves plain=999 / forced=5. The engine now runs in a container
by **default** (image `euro-office-lo:1`): `build_runner()` selects the
container when docker and the image are both available, and otherwise falls back to host
soffice **visibly**. The soffice command is assembled in one place (`HostRunner`/`DockerRunner`),
so the swap is a command-prefix change only, and the recalc logic is identical on both paths.

## Configuration

Both values are injected from the deployment config; neither is stored here.

| variable | secret? | meaning |
|---|---|---|
| `OFFICE_DS_URL` | no | Document Server base URL |
| `OFFICE_DS_JWT` | **yes** | HS256 shared secret for request signing |
| `OFFICE_DS_TIMEOUT_MS` | no | per-request timeout, default 30000 |

The JWT secret is read into memory and used only to sign requests. It is never logged, never
included in an error message, and never leaves the host. Obtaining it from the NC connector
config is a separate, ops-side task.

## Tools

- **`office_status`** — probes `CommandService {c:version}` (reachability + whether the
  signature is accepted) and `{c:license}` (whether recalc is possible). Call it first.
- **`xlsx_set_cells`** — write values/formulas into cells of a workbook. Numbers and booleans
  go in typed; strings go in as inline strings, so `sharedStrings.xml` indices cannot drift.
  Rows and cells are kept in ascending order, which some readers require. Overwriting a cell
  clears whatever was there, so a formula never inherits the previous cached number.
- **`xlsx_recalc`** — recalculate through **headless LibreOffice** and write results back. A
  throwaway per-job profile forces recalc-on-load (`OOXMLRecalcMode`/`ODFRecalcMode` = Always);
  without it the stale cached value is returned unchanged. Success is decided by the OUTPUT
  FILE existing, not the exit code (soffice can exit 0 producing nothing); the process is
  killed on timeout, the temp workspace is always cleaned, and failures return one clean
  message. The original file is left untouched on any failure. Runs on the container engine by
  default (auto when docker and the image are available, else a visible fall back to host
  soffice; `OFFICE_ENGINE` / `OFFICE_SOFFICE_BIN` override), and the result names which engine
  ran (`via container euro-office-lo:1` / `via host soffice /usr/bin/soffice`).

## Verified so far

- `node -c` clean; MCP stdio handshake answers `initialize` and lists all three tools.
- JWT signing checked against an independent HMAC computation: correct HS256 header and
  payload, signature matches, output is base64url-clean, and the secret does not appear in
  the token.
- Missing configuration produces one clear message naming the missing variables.
- `xlsx_set_cells`, against a real workbook and verified by reading the file back:
  - number, text and formula cells written correctly;
  - overwriting a cell that held a cached value with a formula removes the old `<v>` — the
    stale-value case this design exists to prevent;
  - cells inserted before existing ones still come out in ascending column order, and rows in
    ascending row order;
  - `calcChain.xml` dropped and `fullCalcOnLoad` set;
  - error paths return one clear message each: unknown sheet (listing the real sheet names),
    malformed cell reference, missing file, cell with neither value nor formula;
  - the same call driven end-to-end through the MCP tool interface, not just the helper.
- `xlsx_recalc`, end-to-end through the MCP `tools/call` (`office_e2e.py`) against a real
  workbook read back from disk: the deliberately-wrong cached `999` comes back recomputed as
  `5`, and a missing-file call returns one clean error (`isError`, no stack trace). Unit tests
  (`test_office_runner.py`, `test_office_recalc.py` — 41 cases) cover the runner error modes
  (exit-0-without-output, non-zero exit, timeout+kill, missing binary, missing/non-xlsx file),
  temp cleanup on every path, the stale-cache regression (999 -> 5), the Host/Docker command
  seam, and the engine-selection precedence including the visible host fallback. Under AUTO
  selection with docker reachable the tool picks the container and says so (`via container
  euro-office-lo:1`); where docker is not reachable it falls back to host, visibly.

## Next increments

1. Ops rollout — done: the container engine (`euro-office-lo:1`) is the default, and the host
   apt-LibreOffice has been removed, so the container is now the **sole** engine. The
   host-soffice fallback is therefore dormant: with no binary it cannot run, so if the container
   is ever unavailable the tool fails loudly with `soffice not found` — the correct end state,
   not an intermediate one. Container availability is therefore a single point of failure for
   recalc: docker stopping, the image being deleted, or the `user` losing docker-group
   membership each takes recalc down until the container is restored. The gate can still run on
   the container via `office_recalc_gate.py --docker <image>`.
2. Session wiring — project-scope registration for a single agent, following the standard
   MCP-provisioning procedure; requires operator sign-off and a session restart to take effect.
3. docx/pptx/pdf export tools — later increments.

