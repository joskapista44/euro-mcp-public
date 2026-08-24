'use strict'

// M1.3 Core Workbook Operations -- WRITES only through the CURRENT in-memory spreadsheet editor.
// Existing workbooks stay on the live co-edit path. There is intentionally no saved-file parser,
// OOXML rewrite path, or DocBuilder fallback here.

const A1_RE = /^\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/

function validateOperation(op) {
  if (!op || typeof op !== 'object') return 'operation must be an object'
  const type = String(op.type || '')
  const allowed = new Set([
    'sheet_create', 'sheet_copy', 'sheet_delete', 'sheet_rename', 'sheet_move',
    'range_clear', 'range_copy', 'range_move',
    'rows_insert', 'rows_delete', 'columns_insert', 'columns_delete',
  ])
  if (!allowed.has(type)) return `unsupported operation type: ${type || '(empty)'}`
  const sheetNeeded = !['sheet_create'].includes(type)
  if (sheetNeeded && (!op.sheet || typeof op.sheet !== 'string')) return `${type} requires sheet`
  if (['range_clear', 'range_copy', 'range_move'].includes(type) && !A1_RE.test(String(op.range || ''))) return `${type} requires rectangular A1 range`
  if (['range_copy', 'range_move'].includes(type) && !A1_RE.test(String(op.destination || ''))) return `${type} requires destination A1 range`
  if (['rows_insert', 'rows_delete'].includes(type) && (!Number.isInteger(op.start) || op.start < 1 || !Number.isInteger(op.count) || op.count < 1)) return `${type} requires positive integer start and count`
  if (['columns_insert', 'columns_delete'].includes(type) && (!Number.isInteger(op.start) || op.start < 1 || !Number.isInteger(op.count) || op.count < 1)) return `${type} requires positive integer start and count`
  if (['sheet_create', 'sheet_copy'].includes(type) && (!op.name || typeof op.name !== 'string')) return `${type} requires name`
  if (type === 'sheet_rename' && (!op.name || typeof op.name !== 'string')) return 'sheet_rename requires name'
  if (type === 'sheet_move' && (!['before', 'after'].includes(op.position) || !op.referenceSheet)) return 'sheet_move requires position=before|after and referenceSheet'
  return null
}

function columnLabel(n) {
  if (!Number.isInteger(n) || n < 1) return null
  let out = ''
  while (n > 0) { n -= 1; out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26) }
  return out
}

// Stringified and executed by ONLYOFFICE callCommand: keep this function self-contained.
function coreWorkbookOpsCommand(operations) {
  function method(obj, name) { return !!obj && typeof obj[name] === 'function' }
  function colLabel(n) { var s = ''; while (n > 0) { n -= 1; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) } return s }
  function sheetByName(name) { return method(Api, 'GetSheet') ? Api.GetSheet(name) : null }
  function unsupported(index, op, api) { return { ok: false, index: index, type: op.type, outcome: 'unsupported', unsupported: api, error: api + ' is unavailable' } }
  function failure(index, op, outcome, error) { return { ok: false, index: index, type: op.type, outcome: outcome, error: error } }
  function success(index, op, extra) {
    var base = { ok: true, index: index, type: op.type, outcome: 'ok' }
    if (extra) for (var k in extra) base[k] = extra[k]
    return base
  }
  function getRange(sheet, address) {
    if (!sheet || !method(sheet, 'GetRange')) return null
    return sheet.GetRange(address)
  }
  function rowRange(start, count) { return String(start) + ':' + String(start + count - 1) }
  function colRange(start, count) { return colLabel(start) + ':' + colLabel(start + count - 1) }

  try {
    var results = []
    for (var i = 0; i < operations.length; i++) {
      var op = operations[i]
      try {
        if (op.type === 'sheet_create') {
          if (!method(Api, 'AddSheet')) { results.push(unsupported(i, op, 'Api.AddSheet')); break }
          if (sheetByName(op.name)) { results.push(failure(i, op, 'already-exists', 'worksheet already exists')); break }
          var created = Api.AddSheet(op.name)
          results.push(success(i, op, { sheet: created && method(created, 'GetName') ? created.GetName() : op.name }))
        } else if (op.type === 'sheet_copy') {
          var source = sheetByName(op.sheet)
          if (!source) { results.push(failure(i, op, 'sheet-not-found', 'source worksheet was not found')); break }
          if (sheetByName(op.name)) { results.push(failure(i, op, 'already-exists', 'destination worksheet already exists')); break }
          // ApiWorksheet.Copy is not uniformly exposed. Prefer it when present; otherwise fail closed.
          if (!method(source, 'Copy')) { results.push(unsupported(i, op, 'ApiWorksheet.Copy')); break }
          var copied = source.Copy()
          if (!copied) { results.push(failure(i, op, 'copy-failed', 'worksheet Copy returned no worksheet')); break }
          if (!method(copied, 'SetName')) { results.push(unsupported(i, op, 'ApiWorksheet.SetName')); break }
          copied.SetName(op.name)
          results.push(success(i, op, { sheet: op.name, sourceSheet: op.sheet }))
        } else if (op.type === 'sheet_delete') {
          var delSheet = sheetByName(op.sheet)
          if (!delSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet was not found')); break }
          if (!method(delSheet, 'Delete')) { results.push(unsupported(i, op, 'ApiWorksheet.Delete')); break }
          var sheets = method(Api, 'GetSheets') ? Api.GetSheets() : null
          if (sheets && sheets.length <= 1) { results.push(failure(i, op, 'last-sheet', 'refusing to delete the last worksheet')); break }
          var deleted = delSheet.Delete()
          if (deleted === false) { results.push(failure(i, op, 'delete-failed', 'worksheet Delete returned false')); break }
          results.push(success(i, op, { sheet: op.sheet }))
        } else if (op.type === 'sheet_rename') {
          var renSheet = sheetByName(op.sheet)
          if (!renSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet was not found')); break }
          if (sheetByName(op.name)) { results.push(failure(i, op, 'already-exists', 'target worksheet name already exists')); break }
          if (!method(renSheet, 'SetName')) { results.push(unsupported(i, op, 'ApiWorksheet.SetName')); break }
          renSheet.SetName(op.name)
          results.push(success(i, op, { sheet: op.name, previousName: op.sheet }))
        } else if (op.type === 'sheet_move') {
          var moveSheet = sheetByName(op.sheet)
          var refSheet = sheetByName(op.referenceSheet)
          if (!moveSheet || !refSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet or reference worksheet was not found')); break }
          if (!method(moveSheet, 'Move')) { results.push(unsupported(i, op, 'ApiWorksheet.Move')); break }
          if (op.position === 'before') moveSheet.Move(refSheet, null)
          else moveSheet.Move(null, refSheet)
          results.push(success(i, op, { sheet: op.sheet, position: op.position, referenceSheet: op.referenceSheet }))
        } else if (op.type === 'range_clear') {
          var clearSheet = sheetByName(op.sheet)
          if (!clearSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet was not found')); break }
          var clearRange = getRange(clearSheet, op.range)
          if (!clearRange) { results.push(failure(i, op, 'range-not-found', 'range could not be resolved')); break }
          if (!method(clearRange, 'Clear')) { results.push(unsupported(i, op, 'ApiRange.Clear')); break }
          var cleared = clearRange.Clear()
          if (cleared === false) { results.push(failure(i, op, 'clear-failed', 'range Clear returned false')); break }
          results.push(success(i, op, { sheet: op.sheet, range: op.range }))
        } else if (op.type === 'range_copy' || op.type === 'range_move') {
          var sourceSheet = sheetByName(op.sheet)
          var destSheet = sheetByName(op.destinationSheet || op.sheet)
          if (!sourceSheet || !destSheet) { results.push(failure(i, op, 'sheet-not-found', 'source or destination worksheet was not found')); break }
          var src = getRange(sourceSheet, op.range)
          var dst = getRange(destSheet, op.destination)
          if (!src || !dst) { results.push(failure(i, op, 'range-not-found', 'source or destination range could not be resolved')); break }
          var m = op.type === 'range_copy' ? 'Copy' : 'Cut'
          if (!method(src, m)) { results.push(unsupported(i, op, 'ApiRange.' + m)); break }
          src[m](dst)
          results.push(success(i, op, { sheet: op.sheet, range: op.range, destinationSheet: op.destinationSheet || op.sheet, destination: op.destination }))
        } else if (op.type === 'rows_insert' || op.type === 'rows_delete') {
          var rowSheet = sheetByName(op.sheet)
          if (!rowSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet was not found')); break }
          var rr = getRange(rowSheet, rowRange(op.start, op.count))
          if (!rr) { results.push(failure(i, op, 'range-not-found', 'row range could not be resolved')); break }
          var rm = op.type === 'rows_insert' ? 'Insert' : 'Delete'
          if (!method(rr, rm)) { results.push(unsupported(i, op, 'ApiRange.' + rm)); break }
          rr[rm](op.type === 'rows_insert' ? 'down' : 'up')
          results.push(success(i, op, { sheet: op.sheet, start: op.start, count: op.count }))
        } else if (op.type === 'columns_insert' || op.type === 'columns_delete') {
          var colSheet = sheetByName(op.sheet)
          if (!colSheet) { results.push(failure(i, op, 'sheet-not-found', 'worksheet was not found')); break }
          var cr = getRange(colSheet, colRange(op.start, op.count))
          if (!cr) { results.push(failure(i, op, 'range-not-found', 'column range could not be resolved')); break }
          var cm = op.type === 'columns_insert' ? 'Insert' : 'Delete'
          if (!method(cr, cm)) { results.push(unsupported(i, op, 'ApiRange.' + cm)); break }
          cr[cm](op.type === 'columns_insert' ? 'right' : 'left')
          results.push(success(i, op, { sheet: op.sheet, start: op.start, count: op.count }))
        } else {
          results.push(failure(i, op, 'unsupported', 'unknown operation type')); break
        }
      } catch (err) {
        results.push(failure(i, op, 'operation-error', String(err && err.message ? err.message : err)))
        break
      }
    }
    var allOk = results.length === operations.length && results.every(function (r) { return r.ok === true })
    return { ok: allOk, outcome: allOk ? 'ok' : 'partial-or-failed', source: 'live-coedit-editor', applied: results.filter(function (r) { return r.ok }).length, requested: operations.length, results: results }
  } catch (err) {
    return { ok: false, outcome: 'core-workbook-ops-error', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err), results: [] }
  }
}

async function executeInFrame(frame, apiHely, operations, callbackTimeoutMs = 15000) {
  const errors = operations.map(validateOperation).filter(Boolean)
  if (errors.length) return { ok: false, outcome: 'invalid-operation', source: 'live-coedit-editor', error: errors[0], results: [] }
  const body = `return (${coreWorkbookOpsCommand.toString()})(${JSON.stringify(operations)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable on the editor object', results: [] })
    let settled = false
    const finish = (value) => { if (!settled) { settled = true; resolve(value) } }
    try {
      const command = new Function(commandBody)
      editor.callCommand(command, false, false, (value) => finish(value === undefined ? { ok: false, outcome: 'empty-callback', source: 'live-coedit-editor', error: 'callCommand callback returned undefined', results: [] } : value))
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-error', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err), results: [] })
    }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'core workbook operations callback did not arrive in time', results: [] }), timeout)
  }), { u: apiHely, timeout: callbackTimeoutMs, commandBody: body })
}

async function executeLive({ url, user, pass, fileId, operations, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000 }) {
  const errors = operations.map(validateOperation).filter(Boolean)
  if (errors.length) return { ok: false, outcome: 'invalid-operation', source: 'live-coedit-editor', error: errors[0], results: [] }
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', source: 'live-coedit-editor', error: loaded.indok, results: [] }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user)
    await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', source: 'live-coedit-editor', error: 'login did not succeed', results: [] }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', source: 'live-coedit-editor', error: 'spreadsheeteditor frame did not open', results: [] }
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable on known editor objects', results: [] }
    const result = await executeInFrame(frame, apiHely, operations, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely }
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { A1_RE, validateOperation, columnLabel, coreWorkbookOpsCommand, executeInFrame, executeLive }
