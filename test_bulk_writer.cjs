'use strict'

const assert = require('assert')
const fs = require('fs')
const { validateWritePayload, bulkWriterCommand } = require('./bulk-writer.cjs')

function makeApi(names = ['Sheet1']) {
  const sheets = new Map()
  for (const name of names) {
    const cells = new Map()
    sheets.set(name, {
      cells,
      GetRange(address) {
        if (!cells.has(address)) cells.set(address, {
          value: undefined, formula: undefined, cleared: false,
          SetValue(v) { this.value = v },
          SetFormula(f) { this.formula = f },
          Clear() { this.cleared = true; this.value = undefined; this.formula = undefined },
        })
        return cells.get(address)
      },
    })
  }
  return { GetSheet(name) { return sheets.get(name) || null }, sheets }
}

function run(sheet, range, values, formulas, api, maxCells = 26000) {
  const old = global.Api
  global.Api = api
  try { return bulkWriterCommand(sheet, range, values, formulas, maxCells) } finally { global.Api = old }
}

{
  const v = validateWritePayload({ range: 'A1:B2', values: [[1, 'x'], [true, null]], formulas: [[null, null], ['=A1*2', null]] })
  assert.equal(v.ok, true)
  assert.equal(v.parsed.cellCount, 4)
}

{
  const v = validateWritePayload({ range: 'A1:B2', values: [[1, 2, 3]] })
  assert.equal(v.ok, false)
  assert.equal(v.outcome, 'dimension-mismatch')
}

{
  const v = validateWritePayload({ range: 'A1:B1', values: [[1, 2]], formulas: [[null, 'SUM(A1:A2)']] })
  assert.equal(v.ok, false)
  assert.equal(v.outcome, 'invalid-formula')
}

{
  const api = makeApi(['Sheet1', 'Sales'])
  const r = run('Sales', 'B2:C3', [[10, 'hello'], [true, null]], [[null, '=B2*2'], [null, null]], api)
  assert.equal(r.ok, true)
  assert.equal(r.writtenValues, 2)
  assert.equal(r.writtenFormulas, 1)
  assert.equal(r.cleared, 1)
  const cells = api.sheets.get('Sales').cells
  assert.equal(cells.get('B2').value, 10)
  assert.equal(cells.get('C2').formula, '=B2*2')
  assert.equal(cells.get('B3').value, true)
  assert.equal(cells.get('C3').cleared, true)
}

{
  const api = makeApi()
  const rows = 100, cols = 100
  const values = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * cols + c))
  const r = run('Sheet1', 'A1:CV100', values, null, api)
  assert.equal(r.ok, true)
  assert.equal(r.cellCount, 10000)
  assert.equal(r.writtenValues, 10000)
  assert.equal(api.sheets.get('Sheet1').cells.get('CV100').value, 9999)
}

{
  const api = makeApi()
  delete api.sheets.get('Sheet1').GetRange('A1').SetFormula
  const r = run('Sheet1', 'A1', [[0]], [['=1+1']], api)
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'unsupported')
}

{
  const api = makeApi()
  const r = run('Missing', 'A1', [[1]], null, api)
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'sheet-not-found')
}

{
  const src = fs.readFileSync('./bulk-writer.cjs', 'utf8')
  assert.match(src, /callCommand/)
  assert.match(src, /live-coedit-editor/)
  assert.doesNotMatch(src, /DocBuilder.*\(/)
  const entry = fs.readFileSync('./euro-mcp-m14.cjs', 'utf8')
  assert.match(entry, /office_write_range/)
  assert.match(entry, /require\('\.\/euro-mcp-m13\.cjs'\)/)
}

console.log('test_bulk_writer: OK')
