# Euro-MCP

An MCP server for editing Office documents (.docx / .xlsx / .pptx) through a Nextcloud
OnlyOffice Document Server, plus a companion `office-mcp` toolkit for the same file formats
without a live document session (build from scratch, edit an existing file on disk, extract
text/sheets, convert to PDF, recalculate formulas headlessly).

Two related but independent servers live in this repository:

- **`euro-mcp.cjs`** -- the primary server. Talks to a running OnlyOffice Document Server
  session (via a co-editing/DocBuilder box) to read and write documents that are open right
  now, so an edit lands the same way a human typing in the browser would produce it.
- **`office-mcp.js`** -- a standalone toolkit for the same three formats when there is no live
  session: build a document from a structured description, edit an existing file directly,
  extract its text/sheets, export to PDF, and recalculate formulas through headless LibreOffice.
  See `office-mcp.README.md` for its own status and tool list.

## Why this exists

Word/Excel/PowerPoint files store more structure than they look like they do -- a sentence you
see as one line of text can be split across several XML runs, a spreadsheet formula can carry a
stale cached value that nothing recomputes until something opens it, and a package that is
well-formed XML can still fail to open. Both servers exist to handle that structure correctly
rather than treating these formats as plain text, and the test suite (`test_*` / `test-*` files)
is built around the specific traps that caused real, silent output corruption during development
-- each one documents the failure it exists to catch.

## Requirements

- Node.js (for `euro-mcp.cjs` / `office-mcp.js` and the `.cjs`/`.js`/`.mjs` test suite)
- Python 3, no third-party packages (for the `office_*.py` modules and their tests -- the
  docx/pptx builders are written by hand specifically because no install step is assumed)
- LibreOffice (`soffice`), for recalculation and PDF export -- either on the host or via the
  provided `office-lo.Dockerfile`. `OFFICE_ENGINE=docker:<image>` runs it in that container
  (`OFFICE_IMAGE` picks the default image name if you don't build your own);
  `OFFICE_ENGINE=host` drives a host `soffice` binary directly (or set `OFFICE_SOFFICE_BIN`).
- For `euro-mcp.cjs` specifically: SSH access to a box that runs a DocBuilder job runner
  (`box-helper.py`), and a Nextcloud instance with OnlyOffice/DocumentServer for the co-editing
  path (`coedit.cjs`). `office-mcp.js` needs neither.

## Install

```
npm install
cp .env.example .env   # fill in the values your deployment needs; see the file's own comments
```

`.env.example` lists only the variables the code in this repository actually reads
(`process.env.EURO_*` / `OFFICE_*`) -- there is no fleet-specific configuration here.

## Getting a private key to the box (`euro-mcp.cjs`, SSH transport only)

`runner.cjs` needs a private key to reach the DocBuilder box over SSH. Set
`EURO_SSH_KEY_PATH` to a plain private-key file and it is read directly -- no other setup
required. (There is a second, optional integration point, `EURO_VAULT_ROOT`, for pulling the
key from an external secret-vault module instead; that only matters if you already have such a
module, and nothing in this repository provides one.)

## Running the tests

```
node --check *.cjs *.js *.mjs   # syntax
node test-tools.cjs             # end-to-end gate, through the MCP tool interface
node test-lib.cjs               # core translator, no Document Server/box/secret needed
python3 -m unittest discover -p 'test_office_*.py'
python3 -m unittest discover -p 'test_box_helper_*.py'
```

Most tests are pure-function or fake-backed (`fake_soffice.py`, `fake-ds.cjs`) and need no
network access or credentials. A few (`test-coedit-*`) point at a deliberately dead local port
by design, to exercise the network-failure paths without touching a real server.

## Project layout

- `euro-mcp.cjs`, `coedit.cjs`, `lib.cjs`, `lib-operations-*.cjs`, `operations/*.cjs`,
  `euro-magok.cjs`, `runner.cjs`, `office-trace.cjs`, `box-helper.py` -- the live-session
  (DocBuilder/co-editing) server and its supporting modules.
- `office-mcp.js`, `office_*.py`, `fake_soffice*.py` -- the standalone offline toolkit and its
  test doubles.
- `capabilities-registry-build.mjs`, `capability-status-model.mjs` -- tooling that reflects and
  classifies which Office API capabilities are actually confirmed working on a given
  DocumentServer build, versus merely assumed.
- `package-consistency.cjs` -- verifies a written .docx/.xlsx/.pptx package is internally
  consistent (every internal reference the OOXML relationships graph promises actually exists in
  the zip) before it is handed back to the caller.
- `test_*.py`, `test-*.cjs`, `test_*.cjs`, `test-*.mjs` -- the test suite, one file per module
  (or per specific defect, where the filename says so).

## Status

This is a working internal toolkit, exported for public reference. Some rough edges you should
expect from that history:

- `euro-mcp.cjs` and `office-mcp.js` overlap in places (both can write to a workbook, for
  example) because they solve the "live session" and "no session" cases separately rather than
  sharing one code path.
- Comments throughout the code carry provenance references (a decision, a bug report, a
  specific measured trap) written for the internal team that built this -- they are left in
  because they explain the *why* behind a piece of logic that would otherwise look arbitrary,
  even where the reference itself (an internal ticket ID) is not resolvable outside the
  original deployment.
- `coedit.cjs`'s Nextcloud password resolution has exactly one source: `<AGENT>_NEXTCLOUD_APP_PASSWORD`
  in `.env` (see `.env.example`). An earlier internal deployment had a second, vault-service-backed
  fallback here; that service is not part of this repository, so it was removed rather than shipped
  as a code path that would never resolve for anyone outside that original deployment.

## License

See `SECURITY.md` for the vulnerability-reporting process. No `LICENSE` file is included yet;
see the packaging notes for the recommended choice pending a decision from the project owner.

## Contributing

See `CONTRIBUTING.md`.
