'use strict'

const assert = require('assert')
const fs = require('fs')
const path = require('path')
const { layoutCommand } = require('./live-layout.cjs')

let failed = 0
function check(name, fn) {
  try { fn(); console.log(`OK    ${name}`) }
  catch (err) { failed++; console.log(`FAIL  ${name}\n      ${err.stack || err}`) }
}

function fakeApi(options = {}) {
  const state = {
    width: options.width == null ? 10 : options.width,
    height: options.height == null ? 15 : options.height,
    hidden: options.hidden == null ? false : options.hidden,
  }
  const range = {
    GetEntireColumn() { return this },
    GetEntireRow() { return this },
    SetColumnWidth(v) { if (!options.ignoreWidthWrite) state.width = v },
    GetColumnWidth: options.noWidthGetter ? undefined : function () { return state.width },
    SetRowHeight(v) { if (!options.ignoreHeightWrite) state.height = v },
    GetRowHeight: options.noHeightGetter ? undefined : function () { return state.height },
    SetHidden(v) { if (!options.ignoreHiddenWrite) state.hidden = v },
    GetHidden: options.noHiddenGetter ? undefined : function () { return state.hidden },
    AutoFit(rows, cols) {
      if (options.autofitNoChange) return
      if (cols) state.width = options.autofitWidth == null ? 22 : options.autofitWidth
      if (rows) state.height = options.autofitHeight == null ? 28 : options.autofitHeight
    },
  }
  return {
    state,
    api: {
      GetSheet(name) {
        if (name !== 'Sheet1') return null
        return { GetRange() { return range } }
      },
    },
  }
}

function withApi(fixture, fn) {
  const old = global.Api
  global.Api = fixture.api
  try { return fn(fixture.state) } finally {
    if (old === undefined) delete global.Api
    else global.Api = old
  }
}

check('column width write + same-command getter -> PASS', () => {
  const f = fakeApi({ width: 9 })
  const r = withApi(f, () => layoutCommand({ type: 'column.width', sheet: 'Sheet1', range: 'B:B', width: 18 }))
  assert.strictEqual(r.ok, true)
  assert.strictEqual(r.verification.status, 'PASS')
  assert.strictEqual(r.verification.actual, 18)
  assert.strictEqual(r.before, 9)
})

check('column width measurable mismatch -> FAIL', () => {
  const f = fakeApi({ width: 9, ignoreWidthWrite: true })
  const r = withApi(f, () => layoutCommand({ type: 'column.width', sheet: 'Sheet1', range: 'B:B', width: 18 }))
  assert.strictEqual(r.ok, false)
  assert.strictEqual(r.verification.status, 'FAIL')
  assert.strictEqual(r.outcome, 'verification-failed')
})

check('column width without getter -> UNKNOWN, never guessed PASS', () => {
  const f = fakeApi({ noWidthGetter: true })
  const r = withApi(f, () => layoutCommand({ type: 'column.width', sheet: 'Sheet1', range: 'B:B', width: 18 }))
  assert.strictEqual(r.ok, false)
  assert.strictEqual(r.verification.status, 'UNKNOWN')
})

check('row height write + live getter -> PASS', () => {
  const f = fakeApi({ height: 14 })
  const r = withApi(f, () => layoutCommand({ type: 'row.height', sheet: 'Sheet1', range: '3:3', height: 27 }))
  assert.strictEqual(r.ok, true)
  assert.strictEqual(r.verification.status, 'PASS')
  assert.strictEqual(r.verification.actual, 27)
})

check('hide row -> PASS and show row -> PASS', () => {
  const f = fakeApi({ hidden: false })
  const hide = withApi(f, () => layoutCommand({ type: 'rows.hidden', sheet: 'Sheet1', range: '3:5', hidden: true }))
  assert.strictEqual(hide.verification.status, 'PASS')
  assert.strictEqual(hide.verification.actual, true)
  const show = withApi(f, () => layoutCommand({ type: 'rows.hidden', sheet: 'Sheet1', range: '3:5', hidden: false }))
  assert.strictEqual(show.verification.status, 'PASS')
  assert.strictEqual(show.verification.actual, false)
})

check('hide column measurable mismatch -> FAIL', () => {
  const f = fakeApi({ hidden: false, ignoreHiddenWrite: true })
  const r = withApi(f, () => layoutCommand({ type: 'columns.hidden', sheet: 'Sheet1', range: 'C:D', hidden: true }))
  assert.strictEqual(r.verification.status, 'FAIL')
})

check('AutoFit columns with measurable dimension change -> PASS', () => {
  const f = fakeApi({ width: 10, autofitWidth: 24 })
  const r = withApi(f, () => layoutCommand({ type: 'autofit.columns', sheet: 'Sheet1', range: 'A:C' }))
  assert.strictEqual(r.ok, true)
  assert.strictEqual(r.verification.status, 'PASS')
  assert.strictEqual(r.before, 10)
  assert.strictEqual(r.verification.actual, 24)
})

check('AutoFit unchanged dimension -> UNKNOWN, not false FAIL or PASS', () => {
  const f = fakeApi({ width: 10, autofitNoChange: true })
  const r = withApi(f, () => layoutCommand({ type: 'autofit.columns', sheet: 'Sheet1', range: 'A:C' }))
  assert.strictEqual(r.ok, false)
  assert.strictEqual(r.verification.status, 'UNKNOWN')
})

check('invalid dimensions fail closed before mutation', () => {
  const f = fakeApi()
  const r = withApi(f, () => layoutCommand({ type: 'row.height', sheet: 'Sheet1', range: '1:1', height: 0 }))
  assert.strictEqual(r.ok, false)
  assert.strictEqual(r.outcome, 'invalid-operation')
})

check('unknown worksheet fails closed', () => {
  const f = fakeApi()
  const r = withApi(f, () => layoutCommand({ type: 'column.width', sheet: 'Missing', range: 'A:A', width: 10 }))
  assert.strictEqual(r.ok, false)
  assert.strictEqual(r.outcome, 'sheet-not-found')
})

check('M4.3 architecture: no saved-file rewrite implementation and structural ops reuse M1.3', () => {
  const layoutSrc = fs.readFileSync(path.join(__dirname, 'live-layout.cjs'), 'utf8')
  const entrySrc = fs.readFileSync(path.join(__dirname, 'euro-mcp-m43.cjs'), 'utf8')
  assert.ok(!/require\(['"](?:fs|adm-zip|jszip|xlsx)['"]\)/i.test(layoutSrc))
  assert.ok(!/box-helper|runner\.cjs|package-consistency/.test(layoutSrc))
  assert.ok(/require\(['"]\.\/workbook-ops\.cjs['"]\)/.test(entrySrc), 'M1.3 workbook-ops must be reused')
  assert.ok(/verification:[\s\S]*status: 'UNKNOWN'/.test(entrySrc), 'structural mutation must not claim unverified PASS')
})

console.log(`\n${failed === 0 ? 'MIND OK' : `${failed} FAILED`}`)
process.exit(failed === 0 ? 0 : 1)
