'use strict'

const LIVE_SOURCE = 'live-coedit-editor'

function fail(outcome, phase, error, extra = {}) {
  return { ok: false, outcome, phase, error, ...extra }
}

function acceptLive(result, phase, capability) {
  if (!result || typeof result !== 'object') return fail('invalid-result', phase, `${capability} returned no structured result`)
  if (result.ok !== true) return fail('phase-failed', phase, `${capability} failed`, { cause: result })
  if (result.source !== LIVE_SOURCE) return fail('non-live-source', phase, `${capability} did not prove a live co-edit source`, { source: result.source ?? null })
  if (result.outcome !== 'ok') return fail('non-ok-outcome', phase, `${capability} returned a non-ok outcome`, { cause: result })
  return { ok: true, outcome: 'ok', phase, capability, source: LIVE_SOURCE, result }
}

function expectedCell(values, formulas, r, c) {
  const formula = formulas && formulas[r] ? formulas[r][c] : null
  if (formula !== null && formula !== undefined) return { formula, value: undefined, blank: false }
  const value = values[r][c]
  if (value === null || value === undefined) return { formula: null, value: undefined, blank: true }
  return { formula: null, value, blank: false }
}

function cellMatches(cell, expected) {
  if (!cell || typeof cell !== 'object') return false
  if (expected.formula) return cell.formula === expected.formula
  if (expected.blank) return cell.dataType === 'blank' || ((cell.value === '' || cell.value === null || cell.value === undefined) && !cell.formula)
  return Object.is(cell.rawValue, expected.value) || Object.is(cell.value, expected.value) || String(cell.displayText) === String(expected.value)
}

function verifyRangeReadback(after, values, formulas = null) {
  const live = acceptLive(after, 'verify', 'M1.2 office_read_range')
  if (!live.ok) return live
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) return fail('invalid-expectation', 'verify', 'expected values must be a non-empty 2D array')
  const rows = values.length
  const columns = values[0].length
  if (after.rows !== rows || after.columns !== columns || !Array.isArray(after.cells) || after.cells.length !== rows) {
    return fail('verify-shape-mismatch', 'verify', 'read-back shape differs from the requested write', { expected: { rows, columns }, actual: { rows: after.rows ?? null, columns: after.columns ?? null } })
  }
  const mismatches = []
  for (let r = 0; r < rows; r++) {
    if (!Array.isArray(after.cells[r]) || after.cells[r].length !== columns) {
      mismatches.push({ row: r + 1, error: 'row-shape-mismatch' })
      continue
    }
    for (let c = 0; c < columns; c++) {
      const expected = expectedCell(values, formulas, r, c)
      if (!cellMatches(after.cells[r][c], expected)) mismatches.push({ row: r + 1, column: c + 1, address: after.cells[r][c] && after.cells[r][c].address || null, expected, actual: after.cells[r][c] || null })
    }
  }
  if (mismatches.length) return fail('verify-mismatch', 'verify', 'live read-back does not match the requested mutation', { mismatchCount: mismatches.length, mismatches: mismatches.slice(0, 25) })
  return { ok: true, outcome: 'verified', phase: 'verify', source: LIVE_SOURCE, rows, columns, cellCount: rows * columns }
}

function verifyEditEnvelope(edit, capability = 'edit') {
  return acceptLive(edit, 'edit', capability)
}

async function runRangeWriteContract({ inspect, read, edit, values, formulas = null }) {
  const inspection = await inspect()
  const i = acceptLive(inspection, 'inspect', 'M1.1 office_inspect_workbook')
  if (!i.ok) return { ...i, contract: 'inspect -> read -> edit -> verify' }

  const before = await read()
  const b = acceptLive(before, 'read', 'M1.2 office_read_range')
  if (!b.ok) return { ...b, contract: 'inspect -> read -> edit -> verify' }

  const edited = await edit()
  const e = verifyEditEnvelope(edited, 'M1.4 office_write_range')
  if (!e.ok) return { ...e, contract: 'inspect -> read -> edit -> verify', before }

  const after = await read()
  const v = verifyRangeReadback(after, values, formulas)
  if (!v.ok) return { ...v, contract: 'inspect -> read -> edit -> verify', before, edit: edited, after }

  return {
    ok: true,
    outcome: 'verified',
    contract: 'inspect -> read -> edit -> verify',
    source: LIVE_SOURCE,
    phases: { inspect: 'pass', read: 'pass', edit: 'pass', verify: 'pass' },
    before,
    edit: edited,
    after,
    verification: v,
  }
}

module.exports = { LIVE_SOURCE, acceptLive, verifyEditEnvelope, verifyRangeReadback, runRangeWriteContract }
