'use strict'

const assert = require('assert')
const { operationCommand } = require('./workbook-ops.cjs')

function makeApi() {
  const sheets = new Map()
  function mkSheet(name) {
    const ranges = new Map()
    const sheet = {
      _name: name,
      _move: null,
      GetName() { return this._name },
      GetRange(address) {
        if (!ranges.has(address)) ranges.set(address, {
          address,
          cleared: false,
          inserted: [],
          deleted: [],
          copiedTo: null,
          cutTo: null,
          Clear() { this.cleared = true; return true },
          Copy(target) { this.copiedTo = target.address },
          Cut(target) { this.cutTo = target.address },
          Insert(shift) { this.inserted.push(shift) },
          Delete(shift) { this.deleted.push(shift) },
        })
        return ranges.get(address)
      },
      SetName(next) { sheets.delete(this._name); this._name = next; sheets.set(next, this) },
      Delete() { sheets.delete(this._name); return true },
      Copy() { const c = mkSheet(this._name + ' (copy)'); sheets.set(c._name, c); return c },
      Move(before, after) { this._move = { before: before ? before._name : null, after: after ? after._name : null } },
    }
    return sheet
  }
  const api = {
    GetSheet(name) { return sheets.get(name) || null },
    GetSheets() { return Array.from(sheets.values()) },
    AddSheet(name) { const sh = mkSheet(name); sheets.set(name, sh); return sh },
  }
  api.AddSheet('Sheet1')
  return { api, sheets }
}

function run(op, api) {
  const old = global.Api
  global.Api = api
  try { return operationCommand(op) } finally { global.Api = old }
}

{
  const { api, sheets } = makeApi()
  assert.equal(run({ type: 'sheet.create', name: 'Sales' }, api).ok, true)
  assert.equal(run({ type: 'sheet.rename', sheet: 'Sales', name: 'Revenue' }, api).ok, true)
  assert.ok(sheets.has('Revenue'))
  assert.equal(run({ type: 'sheet.copy', sheet: 'Revenue', name: 'Revenue Copy' }, api).ok, true)
  assert.ok(sheets.has('Revenue Copy'))
  api.AddSheet('Anchor')
  assert.equal(run({ type: 'sheet.move', sheet: 'Revenue Copy', position: 'before', referenceSheet: 'Anchor' }, api).ok, true)
  assert.deepEqual(sheets.get('Revenue Copy')._move, { before: 'Anchor', after: null })
  assert.equal(run({ type: 'sheet.move', sheet: 'Revenue Copy', position: 'after', referenceSheet: 'Revenue' }, api).ok, true)
  assert.deepEqual(sheets.get('Revenue Copy')._move, { before: null, after: 'Revenue' })
  assert.equal(run({ type: 'sheet.delete', sheet: 'Revenue Copy' }, api).ok, true)
  assert.ok(!sheets.has('Revenue Copy'))
}

{
  const { api } = makeApi()
  api.AddSheet('Keep')
  let r = run({ type: 'range.clear', sheet: 'Sheet1', range: 'A1:B2' }, api)
  assert.equal(r.ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('A1:B2').cleared, true)

  r = run({ type: 'range.copy', sheet: 'Sheet1', range: 'A1:B2', targetRange: 'D1:E2' }, api)
  assert.equal(r.ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('A1:B2').copiedTo, 'D1:E2')

  r = run({ type: 'range.move', sheet: 'Sheet1', range: 'A1:B2', targetRange: 'G1:H2' }, api)
  assert.equal(r.ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('A1:B2').cutTo, 'G1:H2')

  assert.equal(run({ type: 'rows.insert', sheet: 'Sheet1', range: '2:2' }, api).ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('2:2').inserted[0], 'down')
  assert.equal(run({ type: 'rows.delete', sheet: 'Sheet1', range: '3:3' }, api).ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('3:3').deleted[0], 'up')
  assert.equal(run({ type: 'columns.insert', sheet: 'Sheet1', range: 'B:B' }, api).ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('B:B').inserted[0], 'right')
  assert.equal(run({ type: 'columns.delete', sheet: 'Sheet1', range: 'C:C' }, api).ok, true)
  assert.equal(api.GetSheet('Sheet1').GetRange('C:C').deleted[0], 'left')
}

{
  const { api } = makeApi()
  delete api.GetSheet('Sheet1').Copy
  const r = run({ type: 'sheet.copy', sheet: 'Sheet1', name: 'X' }, api)
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'unsupported')
}

{
  const { api } = makeApi()
  const r = run({ type: 'sheet.delete', sheet: 'Sheet1' }, api)
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'last-sheet')
}

{
  const { api } = makeApi()
  const r = run({ type: 'sheet.rename', sheet: 'Missing', name: 'X' }, api)
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'sheet-not-found')
}

console.log('test_workbook_ops: OK')
