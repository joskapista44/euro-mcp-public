'use strict'

const rangeReader = require('./range-reader.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function formulaInspectorCommand(sheetName, rangeAddress, maxCells) {
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
  function parseCell(cell) {
    var m = String(cell || '').replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/)
    if (!m) return null
    var col = 0
    for (var i = 0; i < m[1].length; i++) col = col * 26 + m[1].toUpperCase().charCodeAt(i) - 64
    return { row: Number(m[2]), column: col }
  }
  function parseRange(address) {
    var p = String(address || '').replace(/\$/g, '').split(':')
    if (p.length > 2) return null
    var a = parseCell(p[0]); var b = parseCell(p[1] || p[0])
    if (!a || !b || b.row < a.row || b.column < a.column) return null
    return { start:a, end:b, rows:b.row-a.row+1, columns:b.column-a.column+1 }
  }
  function fail(outcome, error, extra) {
    var x = { ok:false, outcome:outcome, source:'live-coedit-editor', error:error }
    if (extra) for (var k in extra) x[k] = extra[k]
    return x
  }
  function classify(formulaValue, value, text, formulaSupported) {
    if (!formulaSupported) return { formulaStatus:'unknown', formula:null, cellType:null }
    var f = formulaValue === null || formulaValue === undefined ? '' : String(formulaValue)
    if (f.charAt(0) === '=') return { formulaStatus:'formula', formula:f, cellType:'formula' }
    if ((value === null || value === undefined || value === '') && (text === null || text === undefined || text === '')) return { formulaStatus:'constant', formula:null, cellType:'blank' }
    var t = typeof value
    return { formulaStatus:'constant', formula:null, cellType:t === 'number' || t === 'boolean' || t === 'string' ? t : 'unknown' }
  }

  try {
    if (!method(Api, 'GetSheet')) return fail('unsupported', 'Api.GetSheet is unavailable')
    var sheet = null
    try { sheet = Api.GetSheet(sheetName) } catch (_) {}
    if (!sheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet:sheetName })
    if (!method(sheet, 'GetRange')) return fail('unsupported', 'ApiWorksheet.GetRange is unavailable', { sheet:sheetName })

    var parsed = parseRange(rangeAddress)
    if (!parsed) return fail('invalid-range', 'range must be a rectangular A1 address such as A1:H50', { sheet:sheetName, range:rangeAddress })
    var count = parsed.rows * parsed.columns
    if (count > maxCells) return fail('range-too-large', 'requested range exceeds the configured cell limit', { sheet:sheetName, range:rangeAddress, cellCount:count, maxCells:maxCells })

    var range = null
    try { range = sheet.GetRange(rangeAddress) } catch (_) {}
    if (!range) return fail('range-not-found', 'the requested range could not be resolved', { sheet:sheetName, range:rangeAddress })

    var unsupported = []
    function read(name) {
      if (!method(range, name)) { unsupported.push({field:name, reason:'ApiRange.'+name+' is unavailable'}); return null }
      try { return asMatrix(range[name](), parsed.rows, parsed.columns) }
      catch (err) { unsupported.push({field:name, reason:String(err && err.message ? err.message : err)}); return null }
    }

    var formulas = read('GetFormula')
    var values = read('GetValue')
    var texts = read('GetText')
    var formulaSupported = formulas !== null
    var cells = []
    var summary = { formulas:0, constants:0, blanks:0, unknown:0 }

    for (var r = 0; r < parsed.rows; r++) {
      var row = []
      for (var c = 0; c < parsed.columns; c++) {
        var absRow = parsed.start.row + r
        var absCol = parsed.start.column + c
        var address = colLabel(absCol) + absRow
        var value = valueAt(values, r, c)
        var text = valueAt(texts, r, c)
        var cls = classify(valueAt(formulas, r, c), value, text, formulaSupported)
        if (cls.formulaStatus === 'formula') summary.formulas += 1
        else if (cls.formulaStatus === 'unknown') summary.unknown += 1
        else if (cls.cellType === 'blank') summary.blanks += 1
        else summary.constants += 1
        row.push({ address:address, row:absRow, column:absCol, formulaStatus:cls.formulaStatus, formula:cls.formula, calculatedValue:value, displayText:text, cellType:cls.cellType })
      }
      cells.push(row)
    }

    return { ok:true, outcome:'ok', source:'live-coedit-editor', sheet:sheetName, range:String(rangeAddress).toUpperCase(), rows:parsed.rows, columns:parsed.columns, cellCount:count, cells:cells, summary:summary, unsupported:unsupported }
  } catch (err) {
    return fail('formula-inspector-error', String(err && err.message ? err.message : err), { sheet:sheetName, range:rangeAddress })
  }
}

async function inspectFormulaInFrame(frame, apiHely, { sheet, range, maxCells = 26000, callbackTimeoutMs = 15000 }) {
  const parsed = rangeReader.parseA1Range(range)
  if (!parsed) return { ok:false, outcome:'invalid-range', source:LIVE_SOURCE, sheet, range, error:'range must be a rectangular A1 address such as A1:H50' }
  if (parsed.cellCount > maxCells) return { ok:false, outcome:'range-too-large', source:LIVE_SOURCE, sheet, range:parsed.address, cellCount:parsed.cellCount, maxCells, error:'requested range exceeds the configured cell limit' }
  const body = `return (${formulaInspectorCommand.toString()})(${JSON.stringify(sheet)}, ${JSON.stringify(parsed.address)}, ${Number(maxCells)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable on the editor object'})
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try {
      editor.callCommand(new Function(commandBody), false, false, (value) => finish(value === undefined ? {ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'} : value))
    } catch (err) { finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err && err.message ? err.message : err)}) }
    setTimeout(() => finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'formula inspector callCommand callback did not arrive in time'}), timeout)
  }), { u:apiHely, timeout:callbackTimeoutMs, commandBody:body })
}

async function inspectFormulaLive({ url, user, pass, fileId, sheet, range, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000, maxCells = 26000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok:false, outcome:'nem-mert', source:LIVE_SOURCE, error:loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({viewport:{width:1400,height:900}})
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, {waitUntil:'domcontentloaded', timeout:timeoutMs})
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed (browser remained on login page)'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, {waitUntil:'domcontentloaded',timeout:timeoutMs})
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const result = await inspectFormulaInFrame(frame, apiHely, {sheet,range,maxCells,callbackTimeoutMs})
    return {...result, editor:'spreadsheeteditor', apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

module.exports = { formulaInspectorCommand, inspectFormulaInFrame, inspectFormulaLive }
