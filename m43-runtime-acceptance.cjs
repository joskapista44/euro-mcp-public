'use strict'

// M4.3 combined runtime acceptance probe.
// Acceptance authority is the CURRENT in-memory ONLYOFFICE spreadsheet editor.
// Layout operations use their live getters; structural operations are verified by
// M1.2 live range reads around a temporary marker. No saved-file evidence is used.

const { runLayoutInFrame } = require('./live-layout.cjs')
const { runOperationInFrame } = require('./workbook-ops.cjs')
const { readRangeInFrame } = require('./range-reader.cjs')
const { writeBulkInFrame } = require('./bulk-writer.cjs')

function statusOf(result) {
  if (!result) return 'FAIL'
  if (result.verification && result.verification.status) return result.verification.status
  return result.ok ? 'PASS' : 'FAIL'
}

function overall(steps) {
  const statuses = steps.map((s) => s.status)
  if (statuses.includes('FAIL')) return 'FAIL'
  if (statuses.includes('UNKNOWN')) return 'UNKNOWN'
  return 'PASS'
}

function firstCell(read) {
  return read && read.ok && read.cells && read.cells[0] ? read.cells[0][0] : null
}

function markerVerification(read, expectedAddress, marker) {
  const cell = firstCell(read)
  const actual = cell ? cell.rawValue : undefined
  const pass = !!cell && cell.address === expectedAddress && (actual === marker || cell.value === marker || cell.displayText === marker)
  return {
    status: pass ? 'PASS' : 'FAIL',
    expected: { address: expectedAddress, marker },
    actual: cell ? { address: cell.address, rawValue: cell.rawValue, value: cell.value, displayText: cell.displayText } : null,
  }
}

async function runM43AcceptanceInFrame(frame, apiHely, options = {}) {
  const sheet = options.sheet || 'Sheet1'
  const column = options.column || 'H:H'
  const row = options.row || '12:12'
  const autofitColumn = options.autofitColumn || 'I:I'
  const autofitRow = options.autofitRow || '13:13'
  const structuralRow = options.structuralRow || '20:20'
  const structuralColumn = options.structuralColumn || 'J:J'
  const structuralCell = options.structuralCell || 'J20'
  const rowMovedCell = options.rowMovedCell || 'J21'
  const columnMovedCell = options.columnMovedCell || 'K20'
  const marker = options.marker || `M43-${Date.now()}`
  const width = options.width == null ? 24 : options.width
  const height = options.height == null ? 32 : options.height
  const timeoutMs = options.callbackTimeoutMs || 15000
  const steps = []

  async function layout(name, spec) {
    const result = await runLayoutInFrame(frame, apiHely, spec, timeoutMs)
    steps.push({ name, status: statusOf(result), result })
    return result
  }
  async function read(address) {
    return readRangeInFrame(frame, apiHely, { sheet, range: address, maxCells: 4, callbackTimeoutMs: timeoutMs })
  }
  async function mutate(name, operation) {
    const result = await runOperationInFrame(frame, apiHely, operation, timeoutMs)
    if (!result || !result.ok) {
      steps.push({ name, status: 'FAIL', result, verification: { status: 'FAIL', reason: 'live structural mutation failed' } })
      return false
    }
    return true
  }
  async function verifyMarker(name, address) {
    const observation = await read(address)
    const verification = markerVerification(observation, address, marker)
    steps.push({ name, status: verification.status, observation, verification })
    return verification.status === 'PASS'
  }

  await layout('column-width', { type: 'column.width', sheet, range: column, width })
  await layout('row-height', { type: 'row.height', sheet, range: row, height })
  await layout('column-hide', { type: 'columns.hidden', sheet, range: column, hidden: true })
  await layout('column-show', { type: 'columns.hidden', sheet, range: column, hidden: false })
  await layout('row-hide', { type: 'rows.hidden', sheet, range: row, hidden: true })
  await layout('row-show', { type: 'rows.hidden', sheet, range: row, hidden: false })
  await layout('autofit-column', { type: 'autofit.columns', sheet, range: autofitColumn })
  await layout('autofit-row', { type: 'autofit.rows', sheet, range: autofitRow })

  // Preserve the disposable target cell, seed a unique marker, and prove the seed through M1.2.
  const originalRead = await read(structuralCell)
  const original = firstCell(originalRead)
  if (!original) {
    steps.push({ name: 'structural-pre-read', status: 'FAIL', observation: originalRead, verification: { status: 'FAIL', reason: 'M1.2 could not read the structural marker cell before mutation' } })
  } else {
    steps.push({ name: 'structural-pre-read', status: 'PASS', observation: originalRead, verification: { status: 'PASS', expected: 'readable live cell', actual: original.address } })
    const seed = await writeBulkInFrame(frame, apiHely, { sheet, range: structuralCell, values: [[marker]], callbackTimeoutMs: timeoutMs })
    if (!seed || !seed.ok) {
      steps.push({ name: 'marker-seed', status: 'FAIL', result: seed, verification: { status: 'FAIL', reason: 'live marker write failed' } })
    } else {
      await verifyMarker('marker-seed-readback', structuralCell)

      if (await mutate('row-insert-dispatch', { type: 'rows.insert', sheet, range: structuralRow })) {
        await verifyMarker('row-insert', rowMovedCell)
        if (await mutate('row-delete-dispatch', { type: 'rows.delete', sheet, range: structuralRow })) await verifyMarker('row-delete', structuralCell)
      }

      if (await mutate('column-insert-dispatch', { type: 'columns.insert', sheet, range: structuralColumn })) {
        await verifyMarker('column-insert', columnMovedCell)
        if (await mutate('column-delete-dispatch', { type: 'columns.delete', sheet, range: structuralColumn })) await verifyMarker('column-delete', structuralCell)
      }

      // Restore the original J20 content after the structural shape has been restored.
      const restoreFormula = original.formula || null
      const restoreValue = restoreFormula ? null : original.rawValue
      const restored = await writeBulkInFrame(frame, apiHely, {
        sheet,
        range: structuralCell,
        values: [[restoreValue]],
        formulas: restoreFormula ? [[restoreFormula]] : null,
        callbackTimeoutMs: timeoutMs,
      })
      steps.push({ name: 'marker-restore', status: restored && restored.ok ? 'PASS' : 'FAIL', result: restored,
        verification: { status: restored && restored.ok ? 'PASS' : 'FAIL', expected: 'original J20 content restored for disposable acceptance workbook' } })
    }
  }

  return {
    milestone: 'M4.3',
    source: 'live-coedit-editor',
    outcome: overall(steps),
    humanObservationRequired: false,
    sheet,
    targets: { column, row, autofitColumn, autofitRow, structuralRow, structuralColumn, structuralCell, rowMovedCell, columnMovedCell },
    marker,
    steps,
  }
}

async function runM43AcceptanceLive({ url, user, pass, fileId, loadPlaywright, options = {}, timeoutMs = 60000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { milestone: 'M4.3', outcome: 'FAIL', error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { milestone: 'M4.3', outcome: 'FAIL', error: 'login failed' }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { milestone: 'M4.3', outcome: 'FAIL', error: 'spreadsheeteditor frame not found' }
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { milestone: 'M4.3', outcome: 'FAIL', error: 'callCommand is unavailable' }
    const result = await runM43AcceptanceInFrame(frame, apiHely, options)
    return { ...result, editor: 'spreadsheeteditor', apiHely }
  } finally { await browser.close().catch(() => {}) }
}

module.exports = { statusOf, overall, firstCell, markerVerification, runM43AcceptanceInFrame, runM43AcceptanceLive }
