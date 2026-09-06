'use strict'

const assert = require('assert')
const { structureCommand } = require('./live-structure.cjs')

function makeSheet(name, state) {
  return {
    GetName() { return name },
    Move(before, after) {
      const from = state.order.indexOf(name)
      state.order.splice(from, 1)
      const refName = before ? before.GetName() : after.GetName()
      const ref = state.order.indexOf(refName)
      state.order.splice(before ? ref : ref + 1, 0, name)
    },
    GetRange(address) {
      const first = String(address).split(':')[0]
      const range = {
        Merge() { state.merged[address] = address; return true },
        UnMerge() { delete state.merged[address]; return true },
        GetCells() {
          return {
            get MergeArea() {
              const merged = state.merged[address]
              return { GetAddress() { return merged || first } }
            },
          }
        },
      }
      return range
    },
  }
}

function withApi(names, fn) {
  const state = { order: names.slice(), merged: {} }
  const sheets = {}
  names.forEach((n) => { sheets[n] = makeSheet(n, state) })
  global.Api = {
    GetSheet(name) { return sheets[name] || null },
    GetSheets() { return state.order.map((n) => sheets[n]) },
  }
  try { fn(state, sheets) } finally { delete global.Api }
}

withApi(['A', 'B', 'C'], (state) => {
  const r = structureCommand({ type: 'sheet.inspect' })
  assert.equal(r.ok, true)
  assert.deepEqual(r.order, ['A', 'B', 'C'])
  assert.equal(r.verification.status, 'PASS')

  const moved = structureCommand({ type: 'sheet.move', sheet: 'C', referenceSheet: 'A', position: 'before' })
  assert.equal(moved.ok, true)
  assert.equal(moved.verification.status, 'PASS')
  assert.deepEqual(state.order, ['C', 'A', 'B'])

  const movedAfter = structureCommand({ type: 'sheet.move', sheet: 'C', referenceSheet: 'B', position: 'after' })
  assert.equal(movedAfter.ok, true)
  assert.equal(movedAfter.verification.status, 'PASS')
  assert.deepEqual(state.order, ['A', 'B', 'C'])
})

withApi(['Sheet1'], () => {
  const merged = structureCommand({ type: 'range.merge', sheet: 'Sheet1', range: 'C3:D4' })
  assert.equal(merged.ok, true)
  assert.equal(merged.verification.status, 'PASS')
  assert.equal(merged.verification.actual, 'C3:D4')

  const unmerged = structureCommand({ type: 'range.unmerge', sheet: 'Sheet1', range: 'C3:D4' })
  assert.equal(unmerged.ok, true)
  assert.equal(unmerged.verification.status, 'PASS')
  assert.equal(unmerged.verification.actual, 'C3')
})

withApi(['Sheet1'], (_, sheets) => {
  sheets.Sheet1.GetRange = () => ({ Merge() { return true } })
  const r = structureCommand({ type: 'range.merge', sheet: 'Sheet1', range: 'A1:B2' })
  assert.equal(r.ok, true)
  assert.equal(r.verification.status, 'UNKNOWN')
})

withApi(['Sheet1'], (_, sheets) => {
  sheets.Sheet1.GetRange = () => ({})
  const r = structureCommand({ type: 'range.merge', sheet: 'Sheet1', range: 'A1:B2' })
  assert.equal(r.ok, false)
  assert.equal(r.outcome, 'unsupported')
})

console.log('test_live_structure.cjs: OK')
