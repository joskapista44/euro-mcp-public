'use strict'

const assert = require('assert')
const fs = require('fs')
const reader = require('./range-reader.cjs')

let passed = 0
function test(name, fn) {
  try { fn(); passed += 1; console.log(`ok - ${name}`) }
  catch (err) { console.error(`not ok - ${name}`); throw err }
}

function matrix(rows, cols, fn) {
  return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => fn(r, c)))
}

function withApi(api, fn) {
  const old = global.Api
  global.Api = api
  try { return fn() } finally {
    if (old === undefined) delete global.Api
    else global.Api = old
  }
}

test('A1 parser handles normal, absolute and large ranges', () => {
  assert.deepStrictEqual(reader.parseA1Range('A1:C2'), {
    address: 'A1:C2', start: { row: 1, column: 1 }, end: { row: 2, column: 3 }, rows: 2, columns: 3, cellCount: 6,
  })
  assert.strictEqual(reader.parseA1Range('$A$1:$Z$1000').cellCount, 26000)
  assert.strictEqual(reader.parseA1Range('Sheet2!A1:B2'), null)
  assert.strictEqual(reader.parseA1Range('C3:A1'), null)
})

test('column/address helpers cover columns beyond Z', () => {
  assert.strictEqual(reader.columnNumber('AA'), 27)
  assert.strictEqual(reader.columnLabel(27), 'AA')
  assert.strictEqual(reader.cellAddress(15, 28), 'AB15')
})

test('reader targets the requested non-first worksheet and returns a 2D mixed-type matrix', () => {
  let requestedSheet = null
  let bulkCalls = { GetValue: 0, GetValue2: 0, GetText: 0, GetFormula: 0 }
  const values = [['Alice', 42, 84], [true, '', 45292]]
  const raw = [['Alice', '42', '84'], ['TRUE', '', '45292']]
  const texts = [['Alice', '42', '84'], ['TRUE', '', '2024-01-01']]
  const formulas = [['Alice', '42', '=B1*2'], ['TRUE', '', '45292']]
  const formats = { A1: '@', B1: '0', C1: '0', A2: 'General', B2: 'General', C2: 'yyyy-mm-dd' }

  const sheet = {
    GetRange(address) {
      if (address === 'A1:C2') return {
        GetValue() { bulkCalls.GetValue += 1; return values },
        GetValue2() { bulkCalls.GetValue2 += 1; return raw },
        GetText() { bulkCalls.GetText += 1; return texts },
        GetFormula() { bulkCalls.GetFormula += 1; return formulas },
      }
      return { GetNumberFormat() { return formats[address] ?? null } }
    },
  }
  const api = { GetSheet(name) { requestedSheet = name; return name === 'Second Sheet' ? sheet : null } }
  const result = withApi(api, () => reader.rangeReaderCommand('Second Sheet', 'A1:C2', 26000))

  assert.strictEqual(result.ok, true)
  assert.strictEqual(requestedSheet, 'Second Sheet')
  assert.deepStrictEqual(bulkCalls, { GetValue: 1, GetValue2: 1, GetText: 1, GetFormula: 1 })
  assert.strictEqual(result.rows, 2)
  assert.strictEqual(result.columns, 3)
  assert.strictEqual(result.cellCount, 6)
  assert.strictEqual(result.cells[0][0].dataType, 'string')
  assert.strictEqual(result.cells[0][1].dataType, 'number')
  assert.strictEqual(result.cells[0][2].dataType, 'formula')
  assert.strictEqual(result.cells[0][2].formula, '=B1*2')
  assert.strictEqual(result.cells[0][2].value, 84)
  assert.strictEqual(result.cells[1][0].dataType, 'boolean')
  assert.strictEqual(result.cells[1][1].dataType, 'blank')
  assert.strictEqual(result.cells[1][2].displayText, '2024-01-01')
  assert.strictEqual(result.cells[1][2].numberFormat, 'yyyy-mm-dd')
})

test('single-cell scalar getters are normalized to a 2D result', () => {
  const sheet = {
    GetRange(address) {
      if (address !== 'B7') throw new Error('unexpected range')
      return {
        GetValue() { return 123.5 }, GetValue2() { return '123.5' }, GetText() { return '$123.50' },
        GetFormula() { return '123.5' }, GetNumberFormat() { return '$#,##0.00' },
      }
    },
  }
  const result = withApi({ GetSheet: () => sheet }, () => reader.rangeReaderCommand('Data', 'B7', 26000))
  assert.strictEqual(result.ok, true)
  assert.strictEqual(result.cells[0][0].address, 'B7')
  assert.strictEqual(result.cells[0][0].rawValue, '123.5')
  assert.strictEqual(result.cells[0][0].value, 123.5)
  assert.strictEqual(result.cells[0][0].displayText, '$123.50')
  assert.strictEqual(result.cells[0][0].formula, null)
  assert.strictEqual(result.cells[0][0].dataType, 'number')
})

test('unsupported getters are explicit null/unsupported rather than guessed', () => {
  const sheet = { GetRange: () => ({ GetValue: () => 'x' }) }
  const result = withApi({ GetSheet: () => sheet }, () => reader.rangeReaderCommand('Data', 'A1', 26000))
  assert.strictEqual(result.ok, true)
  assert.strictEqual(result.cells[0][0].rawValue, null)
  assert.strictEqual(result.cells[0][0].displayText, null)
  assert.strictEqual(result.cells[0][0].formula, null)
  assert.strictEqual(result.cells[0][0].numberFormat, null)
  assert(result.unsupported.some((x) => x.field === 'GetValue2'))
  assert(result.unsupported.some((x) => x.field === 'GetText'))
  assert(result.unsupported.some((x) => x.field === 'GetFormula'))
  assert(result.unsupported.some((x) => /numberFormat$/.test(x.field)))
})

test('large A1:Z1000 request stays one command and is accepted at the 26k limit', () => {
  const rows = 1000, cols = 26
  const values = matrix(rows, cols, (r, c) => r * cols + c)
  const texts = matrix(rows, cols, (r, c) => String(r * cols + c))
  const emptyFormula = matrix(rows, cols, () => '')
  let bulkGetRangeCalls = 0
  const sheet = {
    GetRange(address) {
      if (address === 'A1:Z1000') {
        bulkGetRangeCalls += 1
        return { GetValue: () => values, GetValue2: () => texts, GetText: () => texts, GetFormula: () => emptyFormula }
      }
      return { GetNumberFormat: () => 'General' }
    },
  }
  const result = withApi({ GetSheet: () => sheet }, () => reader.rangeReaderCommand('Big', 'A1:Z1000', 26000))
  assert.strictEqual(result.ok, true)
  assert.strictEqual(result.cellCount, 26000)
  assert.strictEqual(result.cells.length, 1000)
  assert.strictEqual(result.cells[999].length, 26)
  assert.strictEqual(result.cells[999][25].address, 'Z1000')
  assert.strictEqual(bulkGetRangeCalls, 1)
})

test('range limit fails closed before reading the requested range', () => {
  let rangeTouched = false
  const sheet = { GetRange() { rangeTouched = true; return null } }
  const result = withApi({ GetSheet: () => sheet }, () => reader.rangeReaderCommand('Big', 'A1:Z1001', 26000))
  assert.strictEqual(result.ok, false)
  assert.strictEqual(result.outcome, 'range-too-large')
  assert.strictEqual(rangeTouched, false)
})

test('package entrypoint preserves M1.2 and the co-edit-only boundary through later wrappers', () => {
  const pkg = require('./package.json')
  assert.match(pkg.main, /^euro-mcp-m\d+\.cjs$/)
  const topEntry = fs.readFileSync(`./${pkg.main}`, 'utf8')
  if (pkg.main !== 'euro-mcp-m12.cjs') assert.match(topEntry, /require\('\.\/euro-mcp-m12\.cjs'\)/)
  const entry = fs.readFileSync('./euro-mcp-m12.cjs', 'utf8')
  assert.match(entry, /office_read_range/)
  assert.match(entry, /readRangeLive/)
  assert.match(entry, /require\('\.\/euro-mcp-m11\.cjs'\)/)
  // Boundary documentation is allowed to name forbidden fallbacks; executable imports/calls are not.
  assert.doesNotMatch(entry, /\brequire\([^\n]*(?:runner|box-helper|package-consistency)/)
  assert.doesNotMatch(entry, /\brunJob\s*\(|\brun_builder_script\s*\(|\bexecFileSync\s*\(/)
  const implementation = fs.readFileSync('./range-reader.cjs', 'utf8')
  assert.match(implementation, /callCommand/)
  assert.match(implementation, /spreadsheeteditor/)
  assert.doesNotMatch(implementation, /\brequire\([^\n]*(?:runner|box-helper|package-consistency)/)
  assert.doesNotMatch(implementation, /\brunJob\s*\(|\bexecFileSync\s*\(/)
})

console.log(`${passed} range reader tests passed`)
