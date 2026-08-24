'use strict'

// M1.3 Core Workbook Operations.
// Existing workbooks are modified only through the CURRENT in-memory ONLYOFFICE
// spreadsheet editor session. No DocBuilder and no saved-OOXML fallback lives here.

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
      if (!has(Api, 'AddSheet')) return unsupported(op.type, 'Api.AddSheet is unavailable')
      if (getSheet(op.name)) return fail('already-exists', 'worksheet already exists', { sheet: op.name })
      var created = Api.AddSheet(op.name)
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.name, result: created === undefined ? null : created }
    }

    if (op.type === 'sheet.rename') {
      var renameSheet = getSheet(op.sheet)
      if (!renameSheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (getSheet(op.name)) return fail('already-exists', 'target worksheet name already exists', { sheet: op.name })
      if (!has(renameSheet, 'SetName')) return unsupported(op.type, 'ApiWorksheet.SetName is unavailable')
      renameSheet.SetName(op.name)
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, name: op.name }
    }

    if (op.type === 'sheet.delete') {
      var deleteSheet = getSheet(op.sheet)
      if (!deleteSheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (has(deleteSheet, 'Delete')) deleteSheet.Delete()
      else if (has(Api, 'DeleteSheet')) Api.DeleteSheet(op.sheet)
      else return unsupported(op.type, 'neither ApiWorksheet.Delete nor Api.DeleteSheet is available')
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet }
    }

    if (op.type === 'sheet.copy') {
      var copySheet = getSheet(op.sheet)
      if (!copySheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (getSheet(op.name)) return fail('already-exists', 'target worksheet name already exists', { sheet: op.name })
      var copied = null
      if (has(copySheet, 'Copy')) copied = copySheet.Copy(op.name)
      else if (has(Api, 'CopySheet')) copied = Api.CopySheet(op.sheet, op.name)
      else return unsupported(op.type, 'neither ApiWorksheet.Copy nor Api.CopySheet is available')
      if (!getSheet(op.name) && copied && has(copied, 'SetName')) copied.SetName(op.name)
      if (!getSheet(op.name)) return fail('verification-failed', 'copy call returned without exposing the target worksheet', { sheet: op.sheet, name: op.name })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, name: op.name }
    }

    if (op.type === 'sheet.move') {
      var moveSheet = getSheet(op.sheet)
      if (!moveSheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: op.sheet })
      if (!Number.isInteger(op.index) || op.index < 0) return fail('invalid-index', 'index must be a zero-based non-negative integer')
      if (has(moveSheet, 'SetPosition')) moveSheet.SetPosition(op.index)
      else if (has(moveSheet, 'Move')) moveSheet.Move(op.index)
      else if (has(Api, 'MoveSheet')) Api.MoveSheet(op.sheet, op.index)
      else return unsupported(op.type, 'no live worksheet move API is available')
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, index: op.index }
    }

    if (op.type === 'range.clear') {
      var rc = requireRange(op.sheet, op.range); if (rc.error) return rc.error
      if (has(rc.range, 'Clear')) rc.range.Clear()
      else if (has(rc.range, 'SetValue')) rc.range.SetValue('')
      else return unsupported(op.type, 'neither ApiRange.Clear nor ApiRange.SetValue is available')
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, range: op.range }
    }

    if (op.type === 'range.copy' || op.type === 'range.move') {
      var src = requireRange(op.sheet, op.range); if (src.error) return src.error
      var dstSheetName = op.targetSheet || op.sheet
      var dst = requireRange(dstSheetName, op.targetRange); if (dst.error) return dst.error
      var did = false
      if (op.type === 'range.copy') {
        if (has(src.range, 'Copy')) { src.range.Copy(dst.range); did = true }
        else if (has(src.range, 'CopyTo')) { src.range.CopyTo(dst.range); did = true }
      } else {
        if (has(src.range, 'Move')) { src.range.Move(dst.range); did = true }
        else if (has(src.range, 'MoveTo')) { src.range.MoveTo(dst.range); did = true }
      }
      if (!did) return unsupported(op.type, 'no compatible live ApiRange copy/move method is available')
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op.type, sheet: op.sheet, range: op.range, targetSheet: dstSheetName, targetRange: op.targetRange }
    }

    if (op.type === 'rows.insert' || op.type === 'rows.delete' || op.type === 'columns.insert' || op.type === 'columns.delete') {
      var structural = requireRange(op.sheet, op.range); if (structural.error) return structural.error
      var isInsert = /\.insert$/.test(op.type)
      var isRows = /^rows\./.test(op.type)
      var shift = isRows ? (isInsert ? 'down' : 'up') : (isInsert ? 'right' : 'left')
      var methodName = isInsert ? 'Insert' : 'Delete'
      if (!has(structural.range, methodName)) return unsupported(op.type, 'ApiRange.' + methodName + ' is unavailable')
      try { structural.range[methodName](shift) }
      catch (e) {
        try { structural.range[methodName]() }
        catch (e2) { return fail('operation-error', String(e2 && e2.message ? e2.message : e2), { operation: op.type }) }
      }
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
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', error: 'login failed' }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', error: 'spreadsheeteditor frame not found' }
    const api = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!api) return { ok: false, outcome: 'nincs-api', error: 'callCommand is unavailable on known editor objects' }
    const result = await runOperationInFrame(frame, api, operation, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely: api }
  } finally { await browser.close().catch(() => {}) }
}

module.exports = { operationCommand, runOperationInFrame, runOperationLive }
