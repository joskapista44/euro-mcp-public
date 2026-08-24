'use strict'

// M1.1 Workbook Inspector -- READS workbook structure from the LIVE spreadsheet editor.
// No OOXML download/parsing and no DocBuilder fallback are allowed here: an already-existing
// workbook belongs to the co-editing route and the source of truth is the editor's in-memory Api.

function columnNumber(label) {
  let n = 0
  for (const ch of String(label || '').toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return null
    n = n * 26 + ch.charCodeAt(0) - 64
  }
  return n || null
}

function parseCell(cell) {
  const m = String(cell || '').replace(/\$/g, '').match(/^([A-Za-z]+)(\d+)$/)
  if (!m) return null
  return { row: Number(m[2]), column: columnNumber(m[1]) }
}

function parseA1Range(address) {
  if (!address) return null
  const raw = String(address).split('!').pop().replace(/\$/g, '')
  const parts = raw.split(':')
  if (parts.length > 2) return null
  const start = parseCell(parts[0])
  const end = parseCell(parts[1] || parts[0])
  if (!start || !end) return null
  return {
    address: String(address),
    start,
    end,
    rows: Math.max(0, end.row - start.row + 1),
    columns: Math.max(0, end.column - start.column + 1),
  }
}

function enrichInspection(raw) {
  if (!raw || raw.ok !== true) return raw
  return {
    ...raw,
    sheets: (raw.sheets || []).map((sheet) => ({
      ...sheet,
      dimensions: parseA1Range(sheet.usedRange),
    })),
  }
}

// Stringified and executed by ONLYOFFICE callCommand: keep it self-contained.
function workbookInspectorCommand() {
  function method(obj, name) { return !!obj && typeof obj[name] === 'function' }
  function safe(label, fn, unsupported) {
    try { return fn() } catch (e) {
      unsupported.push({ field: label, reason: String(e && e.message ? e.message : e) })
      return null
    }
  }

  try {
    var unsupported = []
    var sheets = method(Api, 'GetSheets') ? Api.GetSheets() : []
    if (!method(Api, 'GetSheets')) unsupported.push({ field: 'sheets', reason: 'Api.GetSheets is unavailable' })
    var active = method(Api, 'GetActiveSheet') ? Api.GetActiveSheet() : null
    var activeSheet = active && method(active, 'GetName') ? safe('activeSheet', function () { return active.GetName() }, unsupported) : null
    var resultSheets = []

    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i]
      var name = method(sheet, 'GetName') ? safe('sheets[' + i + '].name', function () { return sheet.GetName() }, unsupported) : null
      var visible = method(sheet, 'GetVisible') ? safe('sheets[' + i + '].visible', function () { return sheet.GetVisible() }, unsupported) : null
      if (!method(sheet, 'GetVisible')) unsupported.push({ field: 'sheets[' + i + '].visible', reason: 'ApiWorksheet.GetVisible is unavailable' })

      var used = method(sheet, 'GetUsedRange') ? safe('sheets[' + i + '].usedRange', function () { return sheet.GetUsedRange() }, unsupported) : null
      if (!method(sheet, 'GetUsedRange')) unsupported.push({ field: 'sheets[' + i + '].usedRange', reason: 'ApiWorksheet.GetUsedRange is unavailable' })
      var usedRange = used && method(used, 'GetAddress') ? safe('sheets[' + i + '].usedRange.address', function () { return used.GetAddress() }, unsupported) : null

      var autoFilter = null
      if (method(sheet, 'GetAutoFilter')) {
        var af = safe('sheets[' + i + '].autoFilter', function () { return sheet.GetAutoFilter() }, unsupported)
        if (af) {
          var afRange = method(af, 'GetRange') ? safe('sheets[' + i + '].autoFilter.range', function () { return af.GetRange() }, unsupported) : null
          autoFilter = {
            present: afRange !== null,
            filterMode: method(af, 'GetFilterMode') ? safe('sheets[' + i + '].autoFilter.filterMode', function () { return af.GetFilterMode() }, unsupported) : null,
            range: afRange && method(afRange, 'GetAddress') ? safe('sheets[' + i + '].autoFilter.range.address', function () { return afRange.GetAddress() }, unsupported) : null,
          }
        }
      } else unsupported.push({ field: 'sheets[' + i + '].autoFilter', reason: 'ApiWorksheet.GetAutoFilter is unavailable' })

      // Table collections vary by ONLYOFFICE edition/version. Probe known names; if none exists,
      // return null + an explicit unsupported entry, never a misleading empty array.
      var tables = null
      var getter = null
      var candidates = ['GetTables', 'GetListObjects', 'GetAllTables']
      for (var c = 0; c < candidates.length; c++) if (method(sheet, candidates[c])) { getter = candidates[c]; break }
      if (getter) {
        var rawTables = safe('sheets[' + i + '].tables', function () { return sheet[getter]() }, unsupported)
        if (rawTables !== null) {
          tables = []
          for (var t = 0; t < rawTables.length; t++) {
            var table = rawTables[t]
            var tableName = method(table, 'GetName') ? safe('sheets[' + i + '].tables[' + t + '].name', function () { return table.GetName() }, unsupported) : null
            var tableRange = null
            if (method(table, 'GetRange')) {
              var tr = safe('sheets[' + i + '].tables[' + t + '].range', function () { return table.GetRange() }, unsupported)
              if (tr && method(tr, 'GetAddress')) tableRange = safe('sheets[' + i + '].tables[' + t + '].range.address', function () { return tr.GetAddress() }, unsupported)
            }
            tables.push({ name: tableName, range: tableRange })
          }
        }
      } else unsupported.push({ field: 'sheets[' + i + '].tables', reason: 'no live worksheet table-collection API found' })

      var localNames = null
      if (method(sheet, 'GetDefNames')) {
        var localRaw = safe('sheets[' + i + '].definedNames', function () { return sheet.GetDefNames() }, unsupported)
        if (localRaw !== null) {
          localNames = []
          for (var ln = 0; ln < localRaw.length; ln++) {
            var lno = localRaw[ln]
            localNames.push({
              name: method(lno, 'GetName') ? safe('sheets[' + i + '].definedNames[' + ln + '].name', function () { return lno.GetName() }, unsupported) : null,
              refersTo: method(lno, 'GetRefersTo') ? safe('sheets[' + i + '].definedNames[' + ln + '].refersTo', function () { return lno.GetRefersTo() }, unsupported) : null,
            })
          }
        }
      } else unsupported.push({ field: 'sheets[' + i + '].definedNames', reason: 'ApiWorksheet.GetDefNames is unavailable' })

      resultSheets.push({ index: i, name: name, visible: visible, active: name !== null && name === activeSheet, usedRange: usedRange, autoFilter: autoFilter, tables: tables, definedNames: localNames })
    }

    var freezePanes = null
    if (method(Api, 'GetFreezePanesType')) freezePanes = safe('freezePanes', function () { return Api.GetFreezePanesType() }, unsupported)
    else unsupported.push({ field: 'freezePanes', reason: 'Api.GetFreezePanesType is unavailable' })

    var definedNames = null
    if (method(Api, 'GetDefNames')) {
      var rawNames = safe('definedNames', function () { return Api.GetDefNames() }, unsupported)
      if (rawNames !== null) {
        definedNames = []
        for (var n = 0; n < rawNames.length; n++) {
          var no = rawNames[n]
          definedNames.push({
            name: method(no, 'GetName') ? safe('definedNames[' + n + '].name', function () { return no.GetName() }, unsupported) : null,
            refersTo: method(no, 'GetRefersTo') ? safe('definedNames[' + n + '].refersTo', function () { return no.GetRefersTo() }, unsupported) : null,
          })
        }
      }
    } else unsupported.push({ field: 'definedNames', reason: 'Api.GetDefNames is unavailable' })

    return { ok: true, outcome: 'ok', source: 'live-coedit-editor', activeSheet: activeSheet, sheetCount: resultSheets.length, sheets: resultSheets, freezePanes: freezePanes, definedNames: definedNames, unsupported: unsupported }
  } catch (err) {
    return { ok: false, outcome: 'inspector-hiba', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) }
  }
}

async function inspectWorkbookInFrame(frame, apiHely, timeoutMs = 10000) {
  const body = `return (${workbookInspectorCommand.toString()})();`
  const result = await frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', error: 'callCommand nem erheto el a megadott editor objektumon' })
    let settled = false
    const finish = (value) => { if (!settled) { settled = true; resolve(value) } }
    try {
      const command = new Function(commandBody)
      // Deployed plugin SDK shape: function, isClose, isCalc, callback.
      editor.callCommand(command, false, false, (value) => finish(value === undefined ? { ok: false, outcome: 'ures-callback', error: 'callCommand callback undefined eredmenyt adott' } : value))
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-dobott', error: String(err && err.message ? err.message : err) })
    }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', error: 'a workbook inspector callCommand callback nem hivodott meg idoben' }), timeout)
  }), { u: apiHely, timeout: timeoutMs, commandBody: body })
  return enrichInspection(result)
}

async function inspectWorkbookLive({ url, user, pass, fileId, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 10000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', error: loaded.indok }
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
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', error: 'a bejelentkezes nem ment at (a login-lapon maradtunk)' }

    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', error: 'spreadsheeteditor frame nem jott fel' }
    const api = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!api) return { ok: false, outcome: 'nincs-api', error: 'a callCommand egyik ismert editor objektumon sem erheto el' }
    const result = await inspectWorkbookInFrame(frame, api, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely: api }
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { columnNumber, parseA1Range, enrichInspection, workbookInspectorCommand, inspectWorkbookInFrame, inspectWorkbookLive }
