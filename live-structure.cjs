'use strict'

const LIVE_SOURCE = 'live-coedit-editor'

function structureCommand(spec) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function fail(outcome, error, extra) { return Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error }, extra || {}) }
  function unknown(extra) { return Object.assign({ status: 'UNKNOWN', reason: 'required live postcondition getter is unavailable' }, extra || {}) }
  function normAddress(value) {
    var s = String(value || '').replace(/\$/g, '').toUpperCase()
    var bang = s.lastIndexOf('!')
    return bang >= 0 ? s.slice(bang + 1) : s
  }
  function firstCellAddress(address) { return normAddress(address).split(':')[0] }
  function getSheet(name) {
    if (!has(Api, 'GetSheet')) return null
    try { return Api.GetSheet(name) } catch (_) { return null }
  }
  function sheetNames() {
    if (!has(Api, 'GetSheets')) return null
    var sheets = null
    try { sheets = Api.GetSheets() } catch (_) { return null }
    if (!sheets || typeof sheets.length !== 'number') return null
    var names = []
    for (var i = 0; i < sheets.length; i++) {
      if (!sheets[i] || !has(sheets[i], 'GetName')) return null
      names.push(String(sheets[i].GetName()))
    }
    return names
  }
  function requireRange(sheetName, address) {
    var sh = getSheet(sheetName)
    if (!sh) return { error: fail('sheet-not-found', 'the requested worksheet was not found', { sheet: sheetName }) }
    if (!has(sh, 'GetRange')) return { error: fail('unsupported', 'ApiWorksheet.GetRange is unavailable', { sheet: sheetName }) }
    var range = null
    try { range = sh.GetRange(address) } catch (_) {}
    if (!range) return { error: fail('range-not-found', 'the requested range could not be resolved', { sheet: sheetName, range: address }) }
    return { sheet: sh, range: range }
  }
  function mergeAreaAddress(range) {
    var probe = range
    if (has(range, 'GetCells')) {
      try { probe = range.GetCells(1, 1) || range } catch (_) {}
    }
    var area = null
    try { area = probe && probe.MergeArea } catch (_) {}
    if (!area || !has(area, 'GetAddress')) return null
    try { return area.GetAddress() } catch (_) { return null }
  }

  try {
    if (!spec || typeof spec.type !== 'string') return fail('invalid-operation', 'spec.type is required')

    if (spec.type === 'sheet.inspect') {
      var order = sheetNames()
      if (!order) return fail('unsupported', 'Api.GetSheets/GetName is unavailable')
      return { ok: true, outcome: 'ok', source: LIVE_SOURCE, order, verification: { status: 'PASS', actual: order } }
    }

    if (spec.type === 'sheet.move') {
      if (!spec.sheet || !spec.referenceSheet || (spec.position !== 'before' && spec.position !== 'after')) return fail('invalid-operation', 'sheet.move requires sheet, referenceSheet and position=before|after')
      if (spec.sheet === spec.referenceSheet) return fail('invalid-operation', 'sheet and referenceSheet must differ')
      var moving = getSheet(spec.sheet)
      var reference = getSheet(spec.referenceSheet)
      if (!moving || !reference) return fail('sheet-not-found', 'worksheet or reference worksheet was not found', { sheet: spec.sheet, referenceSheet: spec.referenceSheet })
      if (!has(moving, 'Move')) return fail('unsupported', 'ApiWorksheet.Move is unavailable')
      var beforeOrder = sheetNames()
      try {
        if (spec.position === 'before') moving.Move(reference, null)
        else moving.Move(null, reference)
      } catch (err) { return fail('operation-error', String(err && err.message ? err.message : err)) }
      var afterOrder = sheetNames()
      if (!afterOrder) return { ok: true, outcome: 'ok', source: LIVE_SOURCE, operation: spec.type, sheet: spec.sheet, referenceSheet: spec.referenceSheet, position: spec.position, beforeOrder, verification: unknown({ expected: spec.position + ' ' + spec.referenceSheet }) }
      var a = afterOrder.indexOf(spec.sheet), b = afterOrder.indexOf(spec.referenceSheet)
      var pass = a >= 0 && b >= 0 && (spec.position === 'before' ? a + 1 === b : b + 1 === a)
      if (!pass) return fail('verification-failed', 'sheet order does not match requested move', { beforeOrder, afterOrder, expected: { sheet: spec.sheet, position: spec.position, referenceSheet: spec.referenceSheet } })
      return { ok: true, outcome: 'ok', source: LIVE_SOURCE, operation: spec.type, sheet: spec.sheet, referenceSheet: spec.referenceSheet, position: spec.position, beforeOrder, afterOrder, verification: { status: 'PASS', expected: { sheet: spec.sheet, position: spec.position, referenceSheet: spec.referenceSheet }, actual: afterOrder } }
    }

    if (spec.type === 'range.merge' || spec.type === 'range.unmerge') {
      if (!spec.sheet || !spec.range) return fail('invalid-operation', 'sheet and range are required')
      var rr = requireRange(spec.sheet, spec.range); if (rr.error) return rr.error
      var method = spec.type === 'range.merge' ? 'Merge' : 'UnMerge'
      if (!has(rr.range, method)) return fail('unsupported', 'ApiRange.' + method + ' is unavailable', { operation: spec.type })
      var mutation
      try { mutation = spec.type === 'range.merge' ? rr.range.Merge(false) : rr.range.UnMerge() }
      catch (err) { return fail('operation-error', String(err && err.message ? err.message : err), { operation: spec.type }) }
      if (mutation === false) return fail('operation-error', 'ApiRange.' + method + ' returned false', { operation: spec.type })

      var actualArea = mergeAreaAddress(rr.range)
      if (actualArea == null) return { ok: true, outcome: 'ok', source: LIVE_SOURCE, operation: spec.type, sheet: spec.sheet, range: spec.range, verification: unknown({ expected: spec.type === 'range.merge' ? normAddress(spec.range) : firstCellAddress(spec.range) }) }
      var expectedArea = spec.type === 'range.merge' ? normAddress(spec.range) : firstCellAddress(spec.range)
      var actualNorm = normAddress(actualArea)
      if (actualNorm !== expectedArea) return fail('verification-failed', 'merge area does not match requested postcondition', { operation: spec.type, expected: expectedArea, actual: actualNorm })
      return { ok: true, outcome: 'ok', source: LIVE_SOURCE, operation: spec.type, sheet: spec.sheet, range: spec.range, verification: { status: 'PASS', expected: expectedArea, actual: actualNorm } }
    }

    return fail('invalid-operation', 'unknown structure operation: ' + spec.type)
  } catch (err) {
    return fail('structure-error', String(err && err.message ? err.message : err), { operation: spec && spec.type ? spec.type : null })
  }
}

async function runStructureInFrame(frame, apiHely, spec, timeoutMs = 15000) {
  const body = `return (${structureCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable' })
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try { editor.callCommand(new Function(commandBody), false, (value) => finish(value === undefined ? { ok: false, outcome: 'empty-callback', source: 'live-coedit-editor', error: 'callCommand callback returned undefined' } : value)) }
    catch (err) { finish({ ok: false, outcome: 'callcommand-error', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) }) }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'structure callCommand callback timed out' }), timeout)
  }), { u: apiHely, timeout: timeoutMs, commandBody: body })
}

async function runStructureLive({ url, user, pass, fileId, spec, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', source: LIVE_SOURCE, error: loaded.indok }
  const { chromium } = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } })
    const page = await ctx.newPage()
    await page.goto(`${url}/login`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.fill('#user', user); await page.fill('#password', pass)
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: timeoutMs }).catch(() => null), page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return { ok: false, outcome: 'auth', source: LIVE_SOURCE, error: 'login failed' }
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`, { waitUntil: 'domcontentloaded', timeout: timeoutMs })
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return { ok: false, outcome: 'nem-nyilt-meg', source: LIVE_SOURCE, error: 'spreadsheeteditor frame not found' }
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return { ok: false, outcome: 'nincs-api', source: LIVE_SOURCE, error: 'callCommand is unavailable' }
    const result = await runStructureInFrame(frame, apiHely, spec, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely }
  } finally { await browser.close().catch(() => {}) }
}

module.exports = { structureCommand, runStructureInFrame, runStructureLive }
