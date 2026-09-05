'use strict'

// M4.3 Row/Column Layout Intelligence.
// All mutations and verification happen inside the CURRENT ONLYOFFICE spreadsheet
// editor session. There is intentionally no saved-file or builder fallback here.

function layoutCommand(spec) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function fail(outcome, error, extra) {
    return Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error }, extra || {})
  }
  function unknown(operation, reason, extra) {
    return Object.assign({
      ok: false,
      outcome: 'unknown',
      source: 'live-coedit-editor',
      operation,
      verification: { status: 'UNKNOWN', reason },
    }, extra || {})
  }
  function pass(operation, expected, actual, extra) {
    return Object.assign({
      ok: true,
      outcome: 'ok',
      source: 'live-coedit-editor',
      operation,
      verification: { status: 'PASS', expected, actual },
    }, extra || {})
  }
  function mismatch(operation, expected, actual, extra) {
    return Object.assign(fail('verification-failed', 'live layout read-back does not match the requested state', {
      operation,
      verification: { status: 'FAIL', expected, actual },
    }), extra || {})
  }
  function finitePositive(v) { return typeof v === 'number' && Number.isFinite(v) && v > 0 }
  function near(a, b) { return typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 0.05 }

  try {
    if (!spec || typeof spec.type !== 'string') return fail('invalid-operation', 'type is required')
    if (!spec.sheet || !spec.range) return fail('invalid-operation', 'sheet and range are required', { operation: spec.type })
    if (!has(Api, 'GetSheet')) return unknown(spec.type, 'Api.GetSheet is unavailable')
    const sheet = Api.GetSheet(spec.sheet)
    if (!sheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: spec.sheet, operation: spec.type })
    if (!has(sheet, 'GetRange')) return unknown(spec.type, 'ApiWorksheet.GetRange is unavailable')
    let range = null
    try { range = sheet.GetRange(spec.range) } catch (_) {}
    if (!range) return fail('range-not-found', 'the requested range could not be resolved', { sheet: spec.sheet, range: spec.range, operation: spec.type })

    if (spec.type === 'column.width') {
      if (!finitePositive(spec.width)) return fail('invalid-operation', 'width must be a positive finite number', { operation: spec.type })
      const target = has(range, 'GetEntireColumn') ? range.GetEntireColumn() : range
      if (!target || !has(target, 'SetColumnWidth')) return unknown(spec.type, 'ApiRange.SetColumnWidth is unavailable')
      const canRead = has(target, 'GetColumnWidth')
      const before = canRead ? target.GetColumnWidth() : null
      target.SetColumnWidth(spec.width)
      if (!canRead) return unknown(spec.type, 'ApiRange.GetColumnWidth is unavailable after a successful write', { before, requested: spec.width })
      const after = target.GetColumnWidth()
      if (!near(after, spec.width)) return mismatch(spec.type, spec.width, after, { before, sheet: spec.sheet, range: spec.range })
      return pass(spec.type, spec.width, after, { before, sheet: spec.sheet, range: spec.range })
    }

    if (spec.type === 'row.height') {
      if (!finitePositive(spec.height)) return fail('invalid-operation', 'height must be a positive finite number', { operation: spec.type })
      const target = has(range, 'GetEntireRow') ? range.GetEntireRow() : range
      if (!target || !has(target, 'SetRowHeight')) return unknown(spec.type, 'ApiRange.SetRowHeight is unavailable')
      const canRead = has(target, 'GetRowHeight')
      const before = canRead ? target.GetRowHeight() : null
      target.SetRowHeight(spec.height)
      if (!canRead) return unknown(spec.type, 'ApiRange.GetRowHeight is unavailable after a successful write', { before, requested: spec.height })
      const after = target.GetRowHeight()
      if (!near(after, spec.height)) return mismatch(spec.type, spec.height, after, { before, sheet: spec.sheet, range: spec.range })
      return pass(spec.type, spec.height, after, { before, sheet: spec.sheet, range: spec.range })
    }

    if (spec.type === 'columns.hidden' || spec.type === 'rows.hidden') {
      if (typeof spec.hidden !== 'boolean') return fail('invalid-operation', 'hidden must be boolean', { operation: spec.type })
      const isColumns = spec.type === 'columns.hidden'
      const getterName = isColumns ? 'GetEntireColumn' : 'GetEntireRow'
      const target = has(range, getterName) ? range[getterName]() : range
      if (!target || !has(target, 'SetHidden')) return unknown(spec.type, 'ApiRange.SetHidden is unavailable')
      const canRead = has(target, 'GetHidden')
      const before = canRead ? target.GetHidden() : null
      target.SetHidden(spec.hidden)
      if (!canRead) return unknown(spec.type, 'ApiRange.GetHidden is unavailable after a successful write', { before, requested: spec.hidden })
      const after = target.GetHidden()
      if (after !== spec.hidden) return mismatch(spec.type, spec.hidden, after, { before, sheet: spec.sheet, range: spec.range })
      return pass(spec.type, spec.hidden, after, { before, sheet: spec.sheet, range: spec.range })
    }

    if (spec.type === 'autofit.columns' || spec.type === 'autofit.rows') {
      const isColumns = spec.type === 'autofit.columns'
      const getterName = isColumns ? 'GetEntireColumn' : 'GetEntireRow'
      const measureName = isColumns ? 'GetColumnWidth' : 'GetRowHeight'
      const target = has(range, getterName) ? range[getterName]() : range
      if (!target || !has(target, 'AutoFit')) return unknown(spec.type, 'ApiRange.AutoFit is unavailable')
      const canRead = has(target, measureName)
      const before = canRead ? target[measureName]() : null
      // ONLYOFFICE AutoFit signature: AutoFit(bRows, bCols).
      target.AutoFit(!isColumns, isColumns)
      if (!canRead) return unknown(spec.type, 'layout dimension getter is unavailable after AutoFit', { before })
      const after = target[measureName]()
      if (typeof after !== 'number' || !Number.isFinite(after) || after <= 0) {
        return unknown(spec.type, 'AutoFit completed but the resulting dimension is not measurable', { before, after })
      }
      if (before === after) {
        return unknown(spec.type, 'AutoFit completed but the measured dimension did not change; the range may already have been best-fit', {
          before, after, sheet: spec.sheet, range: spec.range,
        })
      }
      return pass(spec.type, 'dimension-changed-to-best-fit', after, { before, sheet: spec.sheet, range: spec.range })
    }

    return fail('invalid-operation', 'unknown layout operation: ' + spec.type, { operation: spec.type })
  } catch (err) {
    return fail('operation-error', String(err && err.message ? err.message : err), { operation: spec && spec.type ? spec.type : null })
  }
}

async function runLayoutInFrame(frame, apiHely, spec, timeoutMs = 15000) {
  const body = `return (${layoutCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') {
      return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable' })
    }
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try {
      editor.callCommand(new Function(commandBody), false, false, (value) => finish(value === undefined
        ? { ok: false, outcome: 'ures-callback', source: 'live-coedit-editor', error: 'callCommand returned undefined' }
        : value))
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-dobott', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) })
    }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'layout callback timed out' }), timeout)
  }), { u: apiHely, timeout: timeoutMs, commandBody: body })
}

async function runLayoutLive({ url, user, pass, fileId, spec, loadPlaywright, timeoutMs = 60000, callbackTimeoutMs = 15000 }) {
  const loaded = loadPlaywright()
  if (!loaded.ok) return { ok: false, outcome: 'nem-mert', source: 'live-coedit-editor', error: loaded.indok }
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
    const result = await runLayoutInFrame(frame, api, spec, callbackTimeoutMs)
    return { ...result, editor: 'spreadsheeteditor', apiHely: api }
  } finally {
    await browser.close().catch(() => {})
  }
}

module.exports = { layoutCommand, runLayoutInFrame, runLayoutLive }
