# Design note: one office toolkit for every agent

Nothing implemented yet.

## 1. What exists, and the gap

Live tools (all behind the path guard, all JSON-in/JSON-out helpers spawned per
call, so a fix reaches an agent without restarting it):

| tool | does |
|---|---|
| `office_status` | is the Document Server reachable, is recalc available |
| `xlsx_set_cells` | write values/formulas into cells |
| `xlsx_recalc` | recalculate through headless LibreOffice |
| `office_to_pdf` | xlsx/docx/pptx -> PDF, one sheet or the whole workbook |
| `docx_replace_text` | replace text where Word actually stores it (run boundaries) |
| `docx_set_table_cell` | write one table cell by position |

Missing, and asked for: building a docx and a pptx **from scratch**.

## 2. The pptx (and docx) build decision

Measured on this box, not assumed:

- `python-pptx`: **not installed**. `python-docx`: **not installed**. No `pip`
  (`No module named pip`), and no passwordless root - so a normal install is an
  operator/host decision, not something this work can do.
- `lxml`: **installed**. `Pillow`: **installed**. This matters: those are the
  hard (C-extension) dependencies. `python-pptx` itself is pure Python, so a
  vendored wheel into user-site would very likely work - the pypdf precedent did
  exactly that.
- Getting that wheel needs egress to `files.pythonhosted.org`, which is
  **operator-gated** in this fleet. So option (a) cannot start without an
  explicit approval, and I am not going to smuggle it in as a detail.
- Meanwhile the fleet **already builds both formats by hand with the standard
  library**, documented in `build-docx-no-libs` and `build-pptx-no-libs`, and
  the output is ground-truth verified: a pilot agent's docx (2 pages) and pptx (8 slides)
  rendered correctly through our own engine (`capabilities/office-render.md`).

**Recommendation: implement `docx_build` and `pptx_build` on the proven stdlib
path.** Reasons, in order of weight:

1. it needs no new dependency and no egress approval, so it can ship now;
2. the method is already verified in this fleet, on real output, by another agent
   - we would be wrapping something that works, not betting on something new;
3. the failure mode we care about (a file LibreOffice silently refuses) is
   already understood from the sheet-extraction work, and the render check
   (`office_to_pdf` + page/slide count) catches it.

`python-pptx` stays the documented upgrade path for what a template-based builder
is genuinely bad at: charts, complex master layouts, speaker notes at scale. When
someone asks for those, the egress request is made openly, with that reason.

**Deliberate v1 scope limit.** The builders take a small structured input, not a
document model: title, ordered blocks (heading / paragraph / bullet list /
table / image), and for pptx a list of slides with the same block types. Anything
richer (columns, footnotes, animation) is out of v1 and will be refused with a
clear message rather than half-supported. A builder that silently drops a block
is the same class of bug as a text replace that finds nothing.

## 3. One interface, used the same way by everyone

- Naming stays `<format>_<verb>`: `docx_build`, `pptx_build` join the six above.
  No aliases, no per-agent variants.
- Every tool: absolute paths, checked against `OFFICE_ALLOWED_ROOTS` (default
  `/tmp`); JSON contract with `ok` plus a human-readable message; and where an
  engine runs, the answer says WHICH engine ran.
- Every tool reports what it did in countable terms (cells written, occurrences
  replaced, slides built). Zero is reported, never swallowed.
- One existing agent's wiring is the reference: `agents/<name>/.mcp.json` -> `node
  office-mcp.js`, and `enabledMcpjsonServers` is NOT needed (bypass-mode sessions
  load a project `.mcp.json`; the key would not survive a spawn anyway).

## 4. Provisioning plan (access-gated)

1. `python3 provision-office-mcp.py <agent> --check` - read-only, shows the state.
2. `--dry-run` - shows the exact write.
3. apply, then **restart the pilot agent's session**: a new TOOL is only visible after a
   restart, even though helper fixes are not (the helpers are spawned per call).
   That difference has bitten us twice; it belongs in the runbook.
4. The pilot agent confirms from its own session: build a docx and a pptx into `/tmp`, then
   `office_to_pdf` both and check page/slide count - the same render check that
   already caught a broken file once.
5. Rollback: `--rollback <stamp>` plus a restart.

I prepare and verify; the apply and the restart stay with you and the project owner.

## 5. Documentation

New `capabilities/office-toolkit.md`: the full tool list with input shape and one
worked example per tool, the path-guard rule, the "recalc before export" ordering,
and the restart caveat. `capabilities/office-render.md` keeps its render/recalc
focus and gets a cross-link, so the two do not drift into two versions of the
same instructions.

## 6. What I want a decision on before I build

1. stdlib builders now vs. waiting for a python-pptx egress approval (my
   recommendation: stdlib now, upgrade later on a concrete need);
2. the v1 block list above - anything the project owner needs that is missing from it;
3. Choice of pilot agent, with the restart timed by you.
