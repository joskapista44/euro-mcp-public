'use strict'
const assert = require('assert')
const fs = require('fs')
const { parseA1Range, enrichInspection, workbookInspectorCommand, inspectWorkbookInFrame } = require('./workbook-inspector.cjs')

async function main() {
  assert.deepEqual(parseA1Range('Sales!$B$3:$D$12'), {
    address: 'Sales!$B$3:$D$12',
    start: { row: 3, column: 2 },
    end: { row: 12, column: 4 },
    rows: 10,
    columns: 3,
  })
  assert.equal(parseA1Range('not-a-range'), null)

  const enriched = enrichInspection({ ok: true, sheets: [{ name: 'Sales', usedRange: 'A1:D12' }] })
  assert.equal(enriched.sheets[0].dimensions.rows, 12)
  assert.equal(enriched.sheets[0].dimensions.columns, 4)

  const commandSource = workbookInspectorCommand.toString()
  for (const required of ['GetSheets', 'GetActiveSheet', 'GetVisible', 'GetUsedRange', 'GetAutoFilter', 'GetFreezePanesType', 'GetDefNames']) {
    assert.match(commandSource, new RegExp(required))
  }
  assert.match(commandSource, /no live worksheet table-collection API found/)

  let seen = null
  const frame = {
    evaluate: async (fn, arg) => {
      seen = { fn: String(fn), arg }
      return {
        ok: true,
        outcome: 'ok',
        source: 'live-coedit-editor',
        activeSheet: 'Sales',
        sheetCount: 2,
        sheets: [
          { index: 0, name: 'Sales', visible: true, active: true, usedRange: 'A1:D12', tables: null, autoFilter: { present: true, range: 'A1:D12', filterMode: false }, definedNames: [] },
          { index: 1, name: 'Archive', visible: false, active: false, usedRange: 'A1:B2', tables: null, autoFilter: { present: false, range: null, filterMode: null }, definedNames: [] },
        ],
        freezePanes: 'row',
        definedNames: [{ name: 'Revenue', refersTo: 'Sales!$D$2:$D$12' }],
        unsupported: [{ field: 'sheets[0].tables', reason: 'no live worksheet table-collection API found' }],
      }
    },
  }

  const r = await inspectWorkbookInFrame(frame, 'window.Asc.editor', 4321)
  assert.equal(r.ok, true)
  assert.equal(r.source, 'live-coedit-editor')
  assert.equal(r.activeSheet, 'Sales')
  assert.equal(r.sheetCount, 2)
  assert.equal(r.sheets[0].dimensions.rows, 12)
  assert.equal(r.sheets[1].visible, false)
  assert.equal(r.definedNames[0].name, 'Revenue')
  assert.deepEqual(seen.arg.u, 'window.Asc.editor')
  assert.equal(seen.arg.timeout, 4321)
  assert.match(seen.fn, /callCommand/)
  assert.match(seen.fn, /callback-timeout/)

  // The package entrypoint must still deliver the M1.1 tool after later milestones wrap it.
  // M1.2 deliberately chains euro-mcp-m12 -> euro-mcp-m11 -> base server; pin that chain instead
  // of freezing package.main forever to the M1.1 filename.
  const pkg = require('./package.json')
  assert.equal(pkg.main, 'euro-mcp-m12.cjs')
  const topEntry = fs.readFileSync(`./${pkg.main}`, 'utf8')
  assert.match(topEntry, /require\('\.\/euro-mcp-m11\.cjs'\)/)
  const m11Entry = fs.readFileSync('./euro-mcp-m11.cjs', 'utf8')
  assert.match(m11Entry, /office_inspect_workbook/)
  assert.match(m11Entry, /inspectWorkbookLive/)
  assert.doesNotMatch(m11Entry, /runJob|run_builder_script|DocBuilder fallback/)

  console.log('PASS test_workbook_inspector')
}

main().catch((err) => { console.error(err); process.exit(1) })
