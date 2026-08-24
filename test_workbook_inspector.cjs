'use strict'
const assert = require('assert')
const { inspectWorkbookInFrame } = require('./workbook-inspector.cjs')

async function main() {
  let seen = null
  const frame = {
    evaluate: async (fn, arg) => {
      seen = { fn: String(fn), arg }
      // The browser/editor implementation is intentionally not reproduced in unit tests.
      // This fixture proves the Node-side boundary preserves a JSON-shaped live-editor result.
      return {
        ok: true,
        outcome: 'ok',
        source: 'live-coedit-editor',
        activeSheet: 'Sales',
        sheetCount: 2,
        sheets: [
          { index: 0, name: 'Sales', visible: true, active: true, usedRange: { address: 'A1:D12', rows: 12, columns: 4 }, tables: [], autoFilter: { present: true, range: 'A1:D12', filterMode: false }, freezePanes: { present: true, type: 'row' }, definedNames: [] },
          { index: 1, name: 'Archive', visible: false, active: false, usedRange: { address: 'A1:B2', rows: 2, columns: 2 }, tables: [], autoFilter: { present: false, range: null, filterMode: null }, freezePanes: { present: true, type: null }, definedNames: [] },
        ],
        definedNames: [{ name: 'Revenue', ref: 'Sales!$D$2:$D$12' }],
      }
    },
  }

  const r = await inspectWorkbookInFrame(frame, 'window.Asc.editor', 4321)
  assert.equal(r.ok, true)
  assert.equal(r.source, 'live-coedit-editor')
  assert.equal(r.activeSheet, 'Sales')
  assert.equal(r.sheetCount, 2)
  assert.equal(r.sheets[0].usedRange.address, 'A1:D12')
  assert.equal(r.sheets[1].visible, false)
  assert.equal(r.definedNames[0].name, 'Revenue')
  assert.deepEqual(seen.arg, { u: 'window.Asc.editor', timeout: 4321 })
  assert.match(seen.fn, /callCommand/)
  assert.match(seen.fn, /callback-timeout/)

  console.log('PASS test_workbook_inspector')
}

main().catch((err) => { console.error(err); process.exit(1) })
