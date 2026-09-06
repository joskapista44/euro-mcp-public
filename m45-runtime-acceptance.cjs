'use strict'

const { runOperationInFrame } = require('./workbook-ops.cjs')
const { runTableInFrame } = require('./live-tables.cjs')
const { writeBulkInFrame } = require('./bulk-writer.cjs')
const { readRangeInFrame } = require('./range-reader.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function statusOf(step) {
  return step && step.verification && step.verification.status
    ? step.verification.status
    : (step && step.ok === false ? 'FAIL' : (step && step.ok ? 'PASS' : 'UNKNOWN'))
}

function overall(steps) {
  const statuses = steps.map(statusOf)
  if (statuses.includes('FAIL')) return 'FAIL'
  if (statuses.includes('UNKNOWN')) return 'UNKNOWN'
  return 'PASS'
}

function matrixFromRead(read) {
  if (!read || !read.ok || !Array.isArray(read.cells)) return null
  return read.cells.map((row) => row.map((cell) => cell ? cell.rawValue : null))
}

function sameMatrix(actual, expected) {
  return Array.isArray(actual) && actual.length === expected.length && actual.every((row, r) =>
    Array.isArray(row) && row.length === expected[r].length && row.every((value, c) => String(value) === String(expected[r][c]))
  )
}

async function runM45AcceptanceInFrame(frame, apiHely, suffix = Date.now().toString(36)) {
  const sheet = `M45T_${suffix}`.slice(0, 28)
  const values = [['Product','Price'],['Apples',100],['Oranges',150]]
  const steps = []
  const cleanup = []
  let created = false

  try {
    const create = await runOperationInFrame(frame, apiHely, { type:'sheet.create', name:sheet })
    steps.push({ name:'create-sheet', ...create })
    created = !!(create && create.ok)

    if (created) {
      const seed = await writeBulkInFrame(frame, apiHely, {
        sheet,
        range:'A1:B3',
        values,
        callbackTimeoutMs:15000,
      })
      steps.push({ name:'seed-data-write', ...seed })

      if (seed && seed.ok) {
        const read = await readRangeInFrame(frame, apiHely, { sheet, range:'A1:B3', maxCells:6, callbackTimeoutMs:15000 })
        const actual = matrixFromRead(read)
        const seedPass = !!(read && read.ok && sameMatrix(actual, values))
        steps.push({
          name:'seed-data-readback',
          ok:seedPass,
          outcome:seedPass ? 'ok' : 'verification-failed',
          source:LIVE_SOURCE,
          verification:{ status:seedPass ? 'PASS' : 'FAIL', expected:values, actual },
          observation:read,
        })

        if (seedPass) {
          const createTable = await runTableInFrame(frame, apiHely, {
            type:'table.create',
            sheet,
            range:'A1:B3',
            name:`T_${suffix}`.slice(0,24),
          })
          steps.push({ name:'create-table', ...createTable })

          const inspectTable = await runTableInFrame(frame, apiHely, { type:'table.inspect', sheet })
          steps.push({ name:'inspect-table', ...inspectTable })
        }
      }
    }
  } catch (err) {
    steps.push({ name:'acceptance-exception', ok:false, outcome:'exception', source:LIVE_SOURCE, error:String(err && err.message ? err.message : err), verification:{status:'FAIL'} })
  } finally {
    if (created) {
      try {
        const deleted = await runOperationInFrame(frame, apiHely, { type:'sheet.delete', sheet })
        cleanup.push({ name:'delete-sheet', ...deleted, verification:{ status:deleted && deleted.ok ? 'PASS' : 'FAIL' } })
      } catch (err) {
        cleanup.push({ name:'delete-sheet', ok:false, outcome:'exception', source:LIVE_SOURCE, error:String(err && err.message ? err.message : err), verification:{status:'FAIL'} })
      }
    } else {
      cleanup.push({ name:'delete-sheet', ok:true, outcome:'not-created', source:LIVE_SOURCE, verification:{status:'PASS'} })
    }
  }

  const testOutcome = overall(steps)
  const cleanupOutcome = overall(cleanup)
  const outcome = testOutcome === 'PASS' && cleanupOutcome === 'PASS'
    ? 'PASS'
    : (testOutcome === 'FAIL' || cleanupOutcome === 'FAIL' ? 'FAIL' : 'UNKNOWN')

  return { milestone:'M4.5', source:LIVE_SOURCE, outcome, testOutcome, cleanupOutcome, humanObservationRequired:false, sheet, steps, cleanup }
}

module.exports = { statusOf, overall, matrixFromRead, sameMatrix, runM45AcceptanceInFrame }
