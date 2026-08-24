'use strict'

// M1.3 Core Workbook Operations.
// Existing workbooks are modified only through the CURRENT in-memory ONLYOFFICE
// spreadsheet editor session. No DocBuilder and no saved-file rewrite fallback lives here.

function operationCommand(op) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function fail(outcome, error, extra) { return Object.assign({ ok: false, outcome: outcome, source: 'live-coedit-editor', error: error }, extra || {}) }
  function unsupported(operation, detail) { return fail('unsupported', detail, { operation: operation }) }
  function getSheet(name) {
    if (!has(Api, 'GetSheet')) return null
    try { return Api.GetSheet(name) } catch (_) { return null }
  }
  function requireRange(sheetName, address) {
    var sh = getSheet(sheetName)
    if (!sh) return { error: fail('sheet-not-found', 'the requested worksheet was not found', { sheet: sheetName }) }
    if (!has(sh, 'GetRange')) return { error: unsupported(op.type, 'ApiWorksheet.GetRange is unavailable') }
    var range = null
    try { range = sh.GetRange(address) } catch (_) {}
    if (!range) return { error: fail('range-not-found', 'the requested range could not be resolved', { sheet: sheetName, range: address }) }
    return { sheet: sh, range: range }
  }

  try {
    if (!op || typeof op.type !== 'string') return fail('invalid-operation', 'operation.type is required')

    if (op.type === 'sheet.create') {
      if (!op.name) return fail('invalid-operation', 'name is required', { operation: op.type })
      if (!has(Api, 'AddSheet')) return unsupported(op.type, 'Api.AddSheet is unavailable')
      if (getSheet(op.name)) return fail('already-exists', 'worksheet already exists', { sheet: op.name })
      Api.AddSheet(op.name)
      if (!getSheet(op.name)) return fail('verification-failed', 'AddSheet returned without exposing the new worksheet', { sheet: op.name })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.name }
    }

    if (op.type === 'sheet.rename') {
      if (!op.sheet || !op.name) return fail('invalid-operation', 'sheet and name are required', { operation: op.type })
      var renameSheet = getSheet(op.sheet)
      if (!renameSheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (getSheet(op.name)) return fail('already-exists', 'target worksheet name already exists', { sheet: op.name })
      if (!has(renameSheet, 'SetName')) return unsupported(op.type, 'ApiWorksheet.SetName is unavailable')
      renameSheet.SetName(op.name)
      if (!getSheet(op.name)) return fail('verification-failed', 'SetName completed but the target worksheet name is not visible', { sheet: op.sheet, name: op.name })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, name: op.name }
    }

    if (op.type === 'sheet.delete') {
      if (!op.sheet) return fail('invalid-operation', 'sheet is required', { operation: op.type })
      var deleteSheet = getSheet(op.sheet)
      if (!deleteSheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (has(Api, 'GetSheets')) {
        var currentSheets = Api.GetSheets()
        if (currentSheets && currentSheets.length <= 1) return fail('last-sheet', 'refusing to delete the last worksheet', { sheet: op.sheet })
      }
      if (!has(deleteSheet, 'Delete')) return unsupported(op.type, 'ApiWorksheet.Delete is unavailable')
      var deleted = deleteSheet.Delete()
      if (deleted === false || getSheet(op.sheet)) return fail('verification-failed', 'worksheet Delete did not remove the sheet', { sheet: op.sheet })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet }
    }

    if (op.type === 'sheet.copy') {
      if (!op.sheet || !op.name) return fail('invalid-operation', 'sheet and name are required', { operation: op.type })
      var copySheet = getSheet(op.sheet)
      if (!copySheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (getSheet(op.name)) return fail('already-exists', 'target worksheet name already exists', { sheet: op.name })
      // Worksheet copy is version-dependent in the live API. Probe it and fail closed when absent.
      if (!has(copySheet, 'Copy')) return unsupported(op.type, 'ApiWorksheet.Copy is unavailable')
      var copied = copySheet.Copy()
      if (!copied) return fail('copy-failed', 'ApiWorksheet.Copy returned no worksheet', { sheet: op.sheet })
      if (!has(copied, 'SetName')) return unsupported(op.type, 'copied ApiWorksheet.SetName is unavailable')
      copied.SetName(op.name)
      if (!getSheet(op.name)) return fail('verification-failed', 'copy completed but target worksheet is not visible', { sheet: op.sheet, name: op.name })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, name: op.name }
    }

    if (op.type === 'sheet.move') {
      if (!op.sheet || !op.referenceSheet || (op.position !== 'before' && op.position !== 'after')) return fail('invalid-operation', 'sheet.move requires sheet, referenceSheet and position=before|after', { operation: op.type })
      var moveSheet = getSheet(op.sheet)
      var referenceSheet = getSheet(op.referenceSheet)
      if (!moveSheet || !referenceSheet) return fail('sheet-not-found', 'worksheet or reference worksheet was not found', { sheet: op.sheet, referenceSheet: op.referenceSheet })
      if (!has(moveSheet, 'Move')) return unsupported(op.type, 'ApiWorksheet.Move is unavailable')
      if (op.position === 'before') moveSheet.Move(referenceSheet, null)
      else moveSheet.Move(null, referenceSheet)
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, position: op.position, referenceSheet: op.referenceSheet }
    }

    if (op.type === 'range.clear') {
      var rc = requireRange(op.sheet, op.range); if (rc.error) return rc.error
      if (!has(rc.range, 'Clear')) return unsupported(op.type, 'ApiRange.Clear is unavailable')
      var cleared = rc.range.Clear()
      if (cleared === false) return fail('operation-error', 'ApiRange.Clear returned false', { operation: op.type })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, range: op.range }
    }

    if (op.type === 'range.copy' || op.type === 'range.move') {
      var src = requireRange(op.sheet, op.range); if (src.error) return src.error
      var dstSheetName = op.targetSheet || op.sheet
      var dst = requireRange(dstSheetName, op.targetRange); if (dst.error) return dst.error
      var methodName = op.type === 'range.copy' ? 'Copy' : 'Cut'
      if (!has(src.range, methodName)) return unsupported(op.type, 'ApiRange.' + methodName + ' is unavailable')
      src.range[methodName](dst.range)
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, range: op.range, targetSheet: dstSheetName, targetRange: op.targetRange }
    }

    if (op.type === 'rows.insert' || op.type === 'rows.delete' || op.type === 'columns.insert' || op.type === 'columns.delete') {
      var structural = requireRange(op.sheet, op.range); if (structural.error) return structural.error
      var isInsert = /\.insert$/.test(op.type)
      var isRows = /^rows\./.test(op.type)
      var shift = isRows ? (isInsert ? 'down' : 'up') : (isInsert ? 'right' : 'left')
      var structuralMethod = isInsert ? 'Insert' : 'Delete'
      if (!has(structural.range, structuralMethod)) return unsupported(op.type, 'ApiRange.' + structuralMethod + ' is unavailable')
      structural.range[structuralMethod](shift)
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, range: op.range, shift: shift }
    }

    return fail('invalid-operation', 'unknown operation type: ' + op.type, { operation: op.type })
  } catch (err) {
    return fail('operation-error', String(err && err.message ? err.message : err), { operation: op && op.type ? op.type : null })
  }
}

async function runOperationInFrame(frame, apiHely, operation, timeoutMs = 15000) {
  const body = `return (${operationCommand.toString()})(${JSON.stringify(operation)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable' })
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try {
      editor.callCommand(new Function(commandBody), false, false, (value) => finish(value === undefined ? { ok: false, outcome: 'ures-callback', source: 'live-coedit-editor', error: 'callCommand returned undefined' } : value))
    } catch (err) { finish({ ok: false, outcome: 'callcommand-dobott', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) }) }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'workbook operation callback timed out' }), timeout)
  }), { u: apiHely, timeout: timeoutMs, commandBody: body })
}

async function runOperationLive({ url, user, pass, fileId, operation, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', source: 'live-coedit-editor', error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', source: 'live-coedit-editor', error: 'login failed' }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', source: 'live-coedit-editor', error: 'spreadsheeteditor frame not found' }
    const api = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!api) return { ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable on known editor objects' }
    const result = await runOperationInFrame(frame, api, operation, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely: api }
  } finally { await browser.close().catch(() => {}) }
}

module.exports = { operationCommand, runOperationInFrame, runOperationLive }
