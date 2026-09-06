'use strict'

const assert = require('assert')
const { tableCommand } = require('./live-tables.cjs')

function range(address) {
  return { GetAddress: () => address }
}

function withApi(api, fn) {
  const old = global.Api
  global.Api = api
  try { return fn() } finally {
    if (old === undefined) delete global.Api
    else global.Api = old
  }
}

// The command is serialized into the editor; this catches closure leaks.
{
  const serialized = new Function('spec', 'Api', `return (${tableCommand.toString()})(spec)`)
  const sheet = { FormatAsTable: () => true }
  const out = serialized({ type: 'table.create', sheet: 'S', range: 'A1:B3' }, { GetSheet: () => sheet })
  assert.strictEqual(out.ok, true)
  assert.strictEqual(out.verification.status, 'UNKNOWN')
}

// Native ListObject path: creation + same-session identity/range readback is PASS.
withApi({
  GetSheet: () => ({
    AddListObject: () => ({
      GetName: () => 'SalesTable',
      GetDisplayName: () => 'SalesTable',
      GetRange: () => range('$A$1:$B$3')
    }),
    GetListObjects: () => [{
      GetName: () => 'SalesTable',
      GetDisplayName: () => 'SalesTable',
      GetRange: () => range('A1:B3')
    }]
  })
}, () => {
  const out = tableCommand({ type: 'table.create', sheet: 'S', range: 'A1:B3', name: 'SalesTable' })
  assert.strictEqual(out.ok, true)
  assert.strictEqual(out.verification.status, 'PASS')
})

// Runtime-proven fallback must never manufacture PASS without a getter.
withApi({ GetSheet: () => ({ FormatAsTable: () => true }) }, () => {
  const out = tableCommand({ type: 'table.create', sheet: 'S', range: 'A1:B3' })
  assert.strictEqual(out.ok, true)
  assert.strictEqual(out.verification.status, 'UNKNOWN')
})

console.log('test_live_tables.cjs: OK')
