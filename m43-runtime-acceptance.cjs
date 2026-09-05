'use strict'

// M4.3 combined runtime acceptance probe.
// Runs the complete acceptance sequence in ONE already-open ONLYOFFICE editor
// session, so the human observer can confirm that the same live document changes
// without reload. No WebDAV/saved-XLSX/OOXML evidence is used.

const { runLayoutInFrame } = require('./live-layout.cjs')
const { runOperationInFrame } = require('./workbook-ops.cjs')

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

async function runM43AcceptanceInFrame(frame, apiHely, options = {}) {
  const sheet = options.sheet || 'Sheet1'
  const column = options.column || 'H:H'
  const row = options.row || '12:12'
  const autofitColumn = options.autofitColumn || 'I:I'
  const autofitRow = options.autofitRow || '13:13'
  const structuralCell = options.structuralCell || 'J20'
  const width = options.width == null ? 24 : options.width
  const height = options.height == null ? 32 : options.height
  const timeoutMs = options.callbackTimeoutMs || 15000
  const steps = []

  async function layout(name, spec) {
    const result = await runLayoutInFrame(frame, apiHely, spec, timeoutMs)
    steps.push({ name, status: statusOf(result), result })
    return result
  }
  async function structural(name, operation) {
    const result = await runOperationInFrame(frame, apiHely, operation, timeoutMs)
    // Structural mutations are deliberately UNKNOWN unless the operation itself failed.
    const status = result && result.ok ? 'UNKNOWN' : 'FAIL'
    steps.push({ name, status, result, verification: {
      status,
      reason: status === 'UNKNOWN'
        ? 'live mutation succeeded; human live-editor observation is the runtime acceptance evidence for structural movement'
        : 'live structural mutation failed',
    } })
    return result
  }

  await layout('column-width', { type: 'column.width', sheet, range: column, width })
  await layout('row-height', { type: 'row.height', sheet, range: row, height })
  await layout('column-hide', { type: 'columns.hidden', sheet, range: column, hidden: true })
  await layout('column-show', { type: 'columns.hidden', sheet, range: column, hidden: false })
  await layout('row-hide', { type: 'rows.hidden', sheet, range: row, hidden: true })
  await layout('row-show', { type: 'rows.hidden', sheet, range: row, hidden: false })
  await layout('autofit-column', { type: 'autofit.columns', sheet, range: autofitColumn })
  await layout('autofit-row', { type: 'autofit.rows', sheet, range: autofitRow })

  // Insert/delete pairs intentionally restore the sheet shape after visibly moving
  // the target area. Existing M1.3 workbook-ops is reused for the mutation itself.
  await structural('row-insert', { type: 'rows.insert', sheet, range: structuralCell })
  await structural('row-delete', { type: 'rows.delete', sheet, range: structuralCell })
  await structural('column-insert', { type: 'columns.insert', sheet, range: structuralCell })
  await structural('column-delete', { type: 'columns.delete', sheet, range: structuralCell })

  return {
    milestone: 'M4.3',
    source: 'live-coedit-editor',
    outcome: overall(steps),
    humanObservationRequired: true,
    observerQuestion: 'Did the layout and structural changes appear in the already-open ONLYOFFICE editor without reload?',
    sheet,
    targets: { column, row, autofitColumn, autofitRow, structuralCell },
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
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
      page.click('button[type=submit], input[type=submit]'),
    ])
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
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { statusOf, overall, runM43AcceptanceInFrame, runM43AcceptanceLive }
