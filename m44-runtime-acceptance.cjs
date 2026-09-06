'use strict'

const { runOperationInFrame } = require('./workbook-ops.cjs')
const { runStructureInFrame } = require('./live-structure.cjs')
const { writeBulkInFrame } = require('./bulk-writer.cjs')
const { writeCrossSheetInFrame } = require('./cross-sheet-formulas.cjs')
const { readRangeInFrame } = require('./range-reader.cjs')

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

function formulaIntegrity(read, renamedSheet, expectedValue) {
  const cell = firstCell(read)
  const formula = cell && typeof cell.formula === 'string' ? cell.formula : null
  const actualValue = cell ? cell.rawValue : undefined
  const display = cell ? cell.displayText : undefined
  const noRef = !!formula && !/#REF!/i.test(formula)
  const renamed = !!formula && formula.indexOf(renamedSheet) >= 0
  const valuePass = actualValue === expectedValue || Number(actualValue) === Number(expectedValue) || Number(display) === Number(expectedValue)
  const pass = !!cell && noRef && renamed && valuePass
  return {
    status: pass ? 'PASS' : 'FAIL',
    expected: { renamedSheet, value: expectedValue, noRef: true },
    actual: cell ? { address: cell.address, formula, rawValue: actualValue, displayText: display } : null,
  }
}

async function runM44AcceptanceInFrame(frame, apiHely, options = {}) {
  const suffix = options.suffix || String(Date.now()).slice(-8)
  const sourceSheet = options.sourceSheet || `M44S_${suffix}`
  const renamedSheet = options.renamedSheet || `M44R_${suffix}`
  const calcSheet = options.calcSheet || `M44C_${suffix}`
  const sourceCell = options.sourceCell || 'A1'
  const calcCell = options.calcCell || 'A1'
  const mergeRange = options.mergeRange || 'C3:D4'
  const marker = options.marker == null ? 37 : options.marker
  const expectedValue = Number(marker) * 2
  const timeoutMs = options.callbackTimeoutMs || 15000
  const steps = []
  const cleanup = []

  async function op(name, operation) {
    const result = await runOperationInFrame(frame, apiHely, operation, timeoutMs)
    steps.push({ name, status: statusOf(result), result })
    return result
  }
  async function structure(name, spec) {
    const result = await runStructureInFrame(frame, apiHely, spec, timeoutMs)
    steps.push({ name, status: statusOf(result), result })
    return result
  }
  async function cleanupSheet(sheet) {
    const result = await runOperationInFrame(frame, apiHely, { type: 'sheet.delete', sheet }, timeoutMs)
    cleanup.push({ sheet, status: result && result.ok ? 'PASS' : 'FAIL', result })
  }

  let sourceCurrent = sourceSheet
  let sourceCreated = false
  let calcCreated = false
  try {
    const createSource = await op('sheet-create-source', { type: 'sheet.create', name: sourceSheet })
    sourceCreated = !!(createSource && createSource.ok)
    if (!sourceCreated) return finish()

    const createCalc = await op('sheet-create-calc', { type: 'sheet.create', name: calcSheet })
    calcCreated = !!(createCalc && createCalc.ok)
    if (!calcCreated) return finish()

    const seed = await writeBulkInFrame(frame, apiHely, { sheet: sourceSheet, range: sourceCell, values: [[marker]], callbackTimeoutMs: timeoutMs })
    steps.push({ name: 'source-marker-write', status: seed && seed.ok ? 'PASS' : 'FAIL', result: seed })
    if (!seed || !seed.ok) return finish()

    const formula = `='${sourceSheet.replace(/'/g, "''")}'!${sourceCell}*2`
    const written = await writeCrossSheetInFrame(frame, apiHely, { sheet: calcSheet, range: calcCell, formulas: [[formula]], maxCells: 4, callbackTimeoutMs: timeoutMs })
    steps.push({ name: 'cross-sheet-formula-write', status: written && written.ok && written.verified ? 'PASS' : 'FAIL', result: written })
    if (!written || !written.ok || !written.verified) return finish()

    const beforeRenameRead = await readRangeInFrame(frame, apiHely, { sheet: calcSheet, range: calcCell, maxCells: 4, callbackTimeoutMs: timeoutMs })
    const beforeCell = firstCell(beforeRenameRead)
    const beforePass = !!beforeCell && !/#REF!/i.test(beforeCell.formula || '') && (Number(beforeCell.rawValue) === expectedValue || Number(beforeCell.displayText) === expectedValue)
    steps.push({ name: 'formula-before-rename', status: beforePass ? 'PASS' : 'FAIL', observation: beforeRenameRead })
    if (!beforePass) return finish()

    const rename = await op('sheet-rename-source', { type: 'sheet.rename', sheet: sourceSheet, name: renamedSheet })
    if (!rename || !rename.ok) return finish()
    sourceCurrent = renamedSheet

    await frame.waitForTimeout(250)
    const renamedRead = await readRangeInFrame(frame, apiHely, { sheet: calcSheet, range: calcCell, maxCells: 4, callbackTimeoutMs: timeoutMs })
    const renamedVerification = formulaIntegrity(renamedRead, renamedSheet, expectedValue)
    steps.push({ name: 'formula-after-rename', status: renamedVerification.status, observation: renamedRead, verification: renamedVerification })
    if (renamedVerification.status !== 'PASS') return finish()

    const move = await structure('sheet-move-before-calc', { type: 'sheet.move', sheet: renamedSheet, referenceSheet: calcSheet, position: 'before' })
    if (!move || statusOf(move) !== 'PASS') return finish()

    await frame.waitForTimeout(250)
    const movedRead = await readRangeInFrame(frame, apiHely, { sheet: calcSheet, range: calcCell, maxCells: 4, callbackTimeoutMs: timeoutMs })
    const movedVerification = formulaIntegrity(movedRead, renamedSheet, expectedValue)
    steps.push({ name: 'formula-after-move', status: movedVerification.status, observation: movedRead, verification: movedVerification })
    if (movedVerification.status !== 'PASS') return finish()

    await structure('range-merge', { type: 'range.merge', sheet: calcSheet, range: mergeRange })
    await structure('range-unmerge', { type: 'range.unmerge', sheet: calcSheet, range: mergeRange })
  } finally {
    if (calcCreated) await cleanupSheet(calcSheet)
    if (sourceCreated) await cleanupSheet(sourceCurrent)
  }

  return finish()

  function finish() {
    const testOutcome = overall(steps)
    const cleanupOutcome = cleanup.length && cleanup.every((x) => x.status === 'PASS') ? 'PASS' : (cleanup.length ? 'FAIL' : 'UNKNOWN')
    return {
      milestone: 'M4.4',
      source: 'live-coedit-editor',
      outcome: testOutcome === 'PASS' && cleanupOutcome === 'PASS' ? 'PASS' : (testOutcome === 'FAIL' || cleanupOutcome === 'FAIL' ? 'FAIL' : 'UNKNOWN'),
      testOutcome,
      cleanupOutcome,
      humanObservationRequired: false,
      sheets: { sourceSheet, renamedSheet, calcSheet },
      marker,
      expectedValue,
      mergeRange,
      steps,
      cleanup,
    }
  }
}

async function runM44AcceptanceLive({ url, user, pass, fileId, loadPlaywright, options = {}, timeoutMs = 60000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { milestone: 'M4.4', outcome: 'FAIL', error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { milestone: 'M4.4', outcome: 'FAIL', error: 'login failed' }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { milestone: 'M4.4', outcome: 'FAIL', error: 'spreadsheeteditor frame not found' }
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { milestone: 'M4.4', outcome: 'FAIL', error: 'callCommand is unavailable' }
    const result = await runM44AcceptanceInFrame(frame, apiHely, options)
    return { ...result, editor: 'spreadsheeteditor', apiHely }
  } finally { await browser.close().catch(() => {}) }
}

module.exports = { statusOf, overall, firstCell, formulaIntegrity, runM44AcceptanceInFrame, runM44AcceptanceLive }
