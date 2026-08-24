'use strict'

// M1.2 Live Range Reader -- READS cell content from the CURRENT in-memory spreadsheet editor.
// Architectural boundary: no OOXML download/parsing and no DocBuilder fallback. A workbook that
// already exists belongs to the co-editing path, so the live editor Api is the source of truth.

function columnNumber(label) {
  let n = 0
  for (const ch of String(label || '').toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return null
    n = n * 26 + ch.charCodeAt(0) - 64
  }
  return n || null
}

function columnLabel(n) {
  if (!Number.isInteger(n) || n < 1) return null
  let out = ''
  while (n > 0) {
    n -= 1
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26)
  }
  return out
}

function parseCell(cell) {
  const m = String(cell || '').replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/)
  if (!m) return null
  const row = Number(m[2])
  const column = columnNumber(m[1])
  if (!Number.isSafeInteger(row) || row < 1 || !column) return null
  return { row, column }
}

function parseA1Range(address) {
  const raw = String(address || '').trim().replace(/\$/g, '')
  if (!raw || raw.includes('!')) return null
  const parts = raw.split(':')
  if (parts.length > 2) return null
  const start = parseCell(parts[0])
  const end = parseCell(parts[1] || parts[0])
  if (!start || !end || end.row < start.row || end.column < start.column) return null
  const rows = end.row - start.row + 1
  const columns = end.column - start.column + 1
  return { address: raw.toUpperCase(), start, end, rows, columns, cellCount: rows * columns }
}

function cellAddress(row, column) {
  const col = columnLabel(column)
  return col && Number.isInteger(row) && row > 0 ? `${col}${row}` : null
}

// Stringified and executed by ONLYOFFICE callCommand. Keep this function self-contained.
// All range-wide getters are called once. Per-cell calls are limited to NumberFormat because
// ONLYOFFICE returns that field as a scalar/null for a range, which cannot faithfully represent
// mixed formats. This still remains one browser/callCommand round trip for the whole requested range.
function rangeReaderCommand(sheetName, rangeAddress, maxCells) {
  function method(obj, name) { return !!obj && typeof obj[name] === 'function' }
  function asMatrix(value, rows, cols) {
    if (rows === 1 && cols === 1) return [[value]]
    if (Array.isArray(value) && Array.isArray(value[0])) return value
    if (rows === 1 && Array.isArray(value)) return [value]
    return null
  }
  function valueAt(matrix, r, c) {
    return matrix && matrix[r] && c < matrix[r].length ? matrix[r][c] : null
  }
  function colLabel(n) {
    var s = ''
    while (n > 0) { n -= 1; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) }
    return s
  }
  function parseCellLocal(cell) {
    var m = String(cell || '').replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/)
    if (!m) return null
    var col = 0
    for (var i = 0; i < m[1].length; i++) col = col * 26 + m[1].toUpperCase().charCodeAt(i) - 64
    return { row: Number(m[2]), column: col }
  }
  function parseRangeLocal(address) {
    var p = String(address || '').replace(/\$/g, '').split(':')
    if (p.length > 2) return null
    var a = parseCellLocal(p[0]); var b = parseCellLocal(p[1] || p[0])
    if (!a || !b || b.row < a.row || b.column < a.column) return null
    return { start: a, end: b, rows: b.row - a.row + 1, columns: b.column - a.column + 1 }
  }
  function safe(label, fn, unsupported) {
    try { return fn() } catch (err) {
      unsupported.push({ field: label, reason: String(err && err.message ? err.message : err) })
      return null
    }
  }
  function dataType(value, text, formula) {
    if (formula !== null && formula !== undefined && String(formula).charAt(0) === '=') return 'formula'
    if (value === null || value === undefined || (value === '' && (text === '' || text === null || text === undefined))) return 'blank'
    if (typeof value === 'number') return 'number'
    if (typeof value === 'boolean') return 'boolean'
    if (typeof value === 'string') return 'string'
    return 'unknown'
  }

  try {
    var unsupported = []
    if (!method(Api, 'GetSheet')) return { ok: false, outcome: 'unsupported', source: 'live-coedit-editor', error: 'Api.GetSheet is unavailable' }
    var sheet = safe('sheet', function () { return Api.GetSheet(sheetName) }, unsupported)
    if (!sheet) return { ok: false, outcome: 'sheet-not-found', source: 'live-coedit-editor', sheet: sheetName, error: 'the requested worksheet was not found' }
    if (!method(sheet, 'GetRange')) return { ok: false, outcome: 'unsupported', source: 'live-coedit-editor', sheet: sheetName, error: 'ApiWorksheet.GetRange is unavailable' }

    var parsed = parseRangeLocal(rangeAddress)
    if (!parsed) return { ok: false, outcome: 'invalid-range', source: 'live-coedit-editor', sheet: sheetName, range: rangeAddress, error: 'range must be a rectangular A1 address such as A1:H50' }
    var count = parsed.rows * parsed.columns
    if (count > maxCells) return { ok: false, outcome: 'range-too-large', source: 'live-coedit-editor', sheet: sheetName, range: rangeAddress, cellCount: count, maxCells: maxCells, error: 'requested range exceeds the configured cell limit' }

    var range = safe('range', function () { return sheet.GetRange(rangeAddress) }, unsupported)
    if (!range) return { ok: false, outcome: 'range-not-found', source: 'live-coedit-editor', sheet: sheetName, range: rangeAddress, error: 'the requested range could not be resolved' }

    function readMatrix(name) {
      if (!method(range, name)) {
        unsupported.push({ field: name, reason: 'ApiRange.' + name + ' is unavailable' })
        return null
      }
      return asMatrix(safe(name, function () { return range[name]() }, unsupported), parsed.rows, parsed.columns)
    }

    // Semantics follow ONLYOFFICE's live API: GetValue = stored/calculated value,
    // GetValue2 = unformatted raw value, GetText = displayed text, GetFormula = formula text.
    var values = readMatrix('GetValue')
    var rawValues = readMatrix('GetValue2')
    var texts = readMatrix('GetText')
    var formulas = readMatrix('GetFormula')
    var cells = []

    for (var r = 0; r < parsed.rows; r++) {
      var row = []
      for (var c = 0; c < parsed.columns; c++) {
        var absRow = parsed.start.row + r
        var absCol = parsed.start.column + c
        var address = colLabel(absCol) + absRow
        var value = valueAt(values, r, c)
        var raw = valueAt(rawValues, r, c)
        var text = valueAt(texts, r, c)
        var formula = valueAt(formulas, r, c)
        var numberFormat = null
        var single = safe('cells[' + address + '].range', function () { return sheet.GetRange(address) }, unsupported)
        if (single && method(single, 'GetNumberFormat')) numberFormat = safe('cells[' + address + '].numberFormat', function () { return single.GetNumberFormat() }, unsupported)
        else if (single) unsupported.push({ field: 'cells[' + address + '].numberFormat', reason: 'ApiRange.GetNumberFormat is unavailable' })
        row.push({ address: address, row: absRow, column: absCol, rawValue: raw, value: value, displayText: text, formula: formula && String(formula).charAt(0) === '=' ? formula : null, dataType: dataType(value, text, formula), numberFormat: numberFormat })
      }
      cells.push(row)
    }

    return {
      ok: true,
      outcome: 'ok',
      source: 'live-coedit-editor',
      sheet: sheetName,
      range: String(rangeAddress).toUpperCase(),
      rows: parsed.rows,
      columns: parsed.columns,
      cellCount: count,
      cells: cells,
      unsupported: unsupported,
    }
  } catch (err) {
    return { ok: false, outcome: 'range-reader-error', source: 'live-coedit-editor', sheet: sheetName, range: rangeAddress, error: String(err && err.message ? err.message : err) }
  }
}

async function readRangeInFrame(frame, apiHely, { sheet, range, maxCells = 26000, callbackTimeoutMs = 15000 }) {
  const parsed = parseA1Range(range)
  if (!parsed) return { ok: false, outcome: 'invalid-range', source: 'live-coedit-editor', sheet, range, error: 'range must be a rectangular A1 address such as A1:H50' }
  if (parsed.cellCount > maxCells) return { ok: false, outcome: 'range-too-large', source: 'live-coedit-editor', sheet, range: parsed.address, cellCount: parsed.cellCount, maxCells, error: 'requested range exceeds the configured cell limit' }

  const body = `return (${rangeReaderCommand.toString()})(${JSON.stringify(sheet)}, ${JSON.stringify(parsed.address)}, ${Number(maxCells)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable on the editor object' })
    let settled = false
    const finish = (value) => { if (!settled) { settled = true; resolve(value) } }
    try {
      const command = new Function(commandBody)
      editor.callCommand(command, false, false, (value) => finish(value === undefined ? { ok: false, outcome: 'empty-callback', source: 'live-coedit-editor', error: 'callCommand callback returned undefined' } : value))
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-error', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) })
    }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'range reader callCommand callback did not arrive in time' }), timeout)
  }), { u: apiHely, timeout: callbackTimeoutMs, commandBody: body })
}

async function readRangeLive({ url, user, pass, fileId, sheet, range, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000, maxCells = 26000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', source: 'live-coedit-editor', error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user)
    await page.fill('#password', pass)
    await Promise.all([
      page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null),
      page.click('button[type=submit], input[type=submit]'),
    ])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', source: 'live-coedit-editor', error: 'login did not succeed (browser remained on login page)' }

    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', source: 'live-coedit-editor', error: 'spreadsheeteditor frame did not open' }
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable on known editor objects' }
    const result = await readRangeInFrame(frame, apiHely, { sheet, range, maxCells, callbackTimeoutMs })
    return { ...result, editor: 'spreadsheeteditor', apiHely }
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { columnNumber, columnLabel, parseA1Range, cellAddress, rangeReaderCommand, readRangeInFrame, readRangeLive }
