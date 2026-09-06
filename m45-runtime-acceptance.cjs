'use strict'

const { runOperationInFrame } = require('./workbook-ops.cjs')
const { runTableInFrame } = require('./live-tables.cjs')
const { writeBulkInFrame } = require('./bulk-writer.cjs')

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

async function runM45AcceptanceInFrame(frame, apiHely, suffix = Date.now().toString(36)) {
  const sheet = `M45T_${suffix}`.slice(0, 28)
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
        values:[['Product','Price'],['Apples',100],['Oranges',150]],
        callbackTimeoutMs:15000,
      })
      steps.push({ name:'seed-data', ...seed, verification:{ status:seed && seed.ok && seed.verified ? 'PASS' : 'FAIL' } })

      if (seed && seed.ok && seed.verified) {
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

module.exports = { statusOf, overall, runM45AcceptanceInFrame }
