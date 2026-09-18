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
      return unsupported(op.type, 'sheet.create requires the outer spreadsheet editor API and is dispatched by runOperationInFrame')
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
      return unsupported(op.type, 'sheet.copy requires the outer spreadsheet editor API and is dispatched by runOperationInFrame')
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

async function runNativeSheetCreateInFrame(frame, apiHely, operation, timeoutMs = 15000) {
  return frame.evaluate(async ({ u, op, timeout }) => {
    const e = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    const fail = (outcome, error, extra) => Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error, operation: 'sheet.create' }, extra || {})
    if (!e || typeof e.asc_addWorksheet !== 'function') return fail('unsupported', 'spreadsheet editor asc_addWorksheet is unavailable')
    if (!op?.name) return fail('invalid-operation', 'name is required')
    const names = () => {
      if (typeof e.asc_getWorksheetsCount !== 'function' || typeof e.asc_getWorksheetName !== 'function') return null
      const count = e.asc_getWorksheetsCount(), out = []
      if (!Number.isInteger(count)) return null
      for (let i = 0; i < count; i++) out.push(e.asc_getWorksheetName(i))
      return out
    }
    const before = names()
    if (!before) return fail('inventory-unavailable', 'native worksheet inventory is unavailable')
    if (before.includes(op.name)) return fail('already-exists', 'worksheet already exists', { sheet: op.name })
    const state = {
      canEdit: typeof e.canEdit === 'function' ? e.canEdit() : null,
      globalLock: e.collaborativeEditing && typeof e.collaborativeEditing.getGlobalLock === 'function' ? e.collaborativeEditing.getGlobalLock() : null,
      protectedWorkbook: typeof e.asc_isProtectedWorkbook === 'function' ? e.asc_isProtectedWorkbook() : null
    }
    if (state.canEdit === false || state.globalLock === true || state.protectedWorkbook === true) return fail('create-precondition-blocked', 'the editor cannot acquire the worksheet-create lock', { state })
    const errors = []
    const onError = (id, level) => errors.push({ id, level })
    try { if (typeof e.asc_registerCallback === 'function') e.asc_registerCallback('asc_onError', onError) } catch (_) {}
    try {
      let ret
      try { ret = e.asc_addWorksheet(op.name) }
      catch (err) { return fail('operation-error', String(err && err.message || err), { sheet: op.name, state, errors }) }
      if (ret === false) return fail('create-failed', 'asc_addWorksheet returned false', { sheet: op.name, state, errors })
      const started = Date.now()
      while (Date.now() - started <= timeout) {
        const current = names()
        if (current) {
          const hits = current.filter(name => name === op.name).length
          if (hits === 1) return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: 'sheet.create', sheet: op.name, beforeCount: before.length, afterCount: current.length, dispatchApi: 'asc_addWorksheet', state, errors }
          if (hits > 1) return fail('verification-failed', 'target worksheet identity became ambiguous', { sheet: op.name, targetCount: hits, state, errors })
        }
        await new Promise(resolve => setTimeout(resolve, 25))
      }
      return fail('collaboration-lock-timeout', 'asc_addWorksheet did not expose the target worksheet before the state deadline', { sheet: op.name, state, errors })
    } finally {
      try { if (typeof e.asc_unregisterCallback === 'function') e.asc_unregisterCallback('asc_onError', onError) } catch (_) {}
    }
  }, { u: apiHely, op: operation, timeout: timeoutMs })
}

async function runNativeSheetMoveInFrame(frame, apiHely, operation, timeoutMs = 15000) {
  return frame.evaluate(async ({ u, op, timeout }) => {
    const e = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    const fail = (outcome, error, extra) => Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error, operation: 'sheet.move' }, extra || {})
    if (!e || typeof e.asc_moveWorksheet !== 'function') return fail('unsupported', 'spreadsheet editor asc_moveWorksheet is unavailable')
    if (!op?.sheet || !op?.referenceSheet || !['before','after'].includes(op.position)) return fail('invalid-operation', 'sheet, referenceSheet and position=before|after are required')
    const names = () => {
      if (typeof e.asc_getWorksheetsCount !== 'function' || typeof e.asc_getWorksheetName !== 'function') return null
      const count = e.asc_getWorksheetsCount(), out = []
      if (!Number.isInteger(count)) return null
      for (let i = 0; i < count; i++) out.push(e.asc_getWorksheetName(i))
      return out
    }
    const before = names()
    if (!before) return fail('inventory-unavailable', 'native worksheet inventory is unavailable')
    const sourceIndex = before.indexOf(op.sheet), referenceIndex = before.indexOf(op.referenceSheet)
    if (sourceIndex < 0 || referenceIndex < 0 || sourceIndex === referenceIndex) return fail('sheet-not-found', 'worksheet or reference worksheet was not found uniquely', { sheet: op.sheet, referenceSheet: op.referenceSheet })
    const state = {
      canEdit: typeof e.canEdit === 'function' ? e.canEdit() : null,
      globalLock: e.collaborativeEditing && typeof e.collaborativeEditing.getGlobalLock === 'function' ? e.collaborativeEditing.getGlobalLock() : null,
      protectedWorkbook: typeof e.asc_isProtectedWorkbook === 'function' ? e.asc_isProtectedWorkbook() : null
    }
    if (state.canEdit === false || state.globalLock === true || state.protectedWorkbook === true) return fail('move-precondition-blocked', 'the editor cannot acquire the worksheet-move lock', { state })
    const errors = [], onError = (id, level) => errors.push({ id, level })
    try { if (typeof e.asc_registerCallback === 'function') e.asc_registerCallback('asc_onError', onError) } catch (_) {}
    try {
      const where = op.position === 'before' ? referenceIndex : referenceIndex + 1
      let ret
      try { ret = e.asc_moveWorksheet(where, [sourceIndex]) }
      catch (err) { return fail('operation-error', String(err && err.message || err), { sheet: op.sheet, referenceSheet: op.referenceSheet, position: op.position, state, errors }) }
      if (ret === false) return fail('move-failed', 'asc_moveWorksheet returned false', { sheet: op.sheet, referenceSheet: op.referenceSheet, position: op.position, state, errors })
      const started = Date.now()
      while (Date.now() - started <= timeout) {
        const current = names()
        if (current) {
          const si = current.indexOf(op.sheet), ri = current.indexOf(op.referenceSheet)
          const matched = op.position === 'before' ? si === ri - 1 : si === ri + 1
          if (matched) return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: 'sheet.move', sheet: op.sheet, referenceSheet: op.referenceSheet, position: op.position, beforeIndex: sourceIndex, afterIndex: si, referenceIndex: ri, dispatchApi: 'asc_moveWorksheet', state, errors }
        }
        await new Promise(resolve => setTimeout(resolve, 25))
      }
      return fail('collaboration-lock-timeout', 'asc_moveWorksheet did not expose the requested order before the state deadline', { sheet: op.sheet, referenceSheet: op.referenceSheet, position: op.position, state, errors })
    } finally {
      try { if (typeof e.asc_unregisterCallback === 'function') e.asc_unregisterCallback('asc_onError', onError) } catch (_) {}
    }
  }, { u: apiHely, op: operation, timeout: timeoutMs })
}

async function runNativeSheetCopyInFrame(frame, apiHely, operation, timeoutMs = 15000) {
  return frame.evaluate(async ({ u, op, timeout }) => {
    const e = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    const fail = (outcome, error, extra) => Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error, operation: 'sheet.copy' }, extra || {})
    if (!e || typeof e.asc_copyWorksheet !== 'function') return fail('unsupported', 'spreadsheet editor asc_copyWorksheet is unavailable')
    if (!op?.sheet || !op?.name) return fail('invalid-operation', 'sheet and name are required')
    const inventory = await new Promise((resolve) => {
      let settled = false
      const finish = value => { if (!settled) { settled = true; resolve(value) } }
      try {
        e.callCommand(new Function(`try{if(typeof Api==='undefined'||!Api||typeof Api.GetSheets!=='function')return {ok:false,outcome:'unsupported',error:'Api.GetSheets is unavailable'};var a=Api.GetSheets()||[],names=[];for(var i=0;i<a.length;i++){var s=a[i];names.push(s&&typeof s.GetName==='function'?s.GetName():null)}return {ok:true,names:names}}catch(err){return {ok:false,outcome:'operation-error',error:String(err&&err.message?err.message:err)}}`), false, value => finish(value))
      } catch (err) { finish({ ok:false, outcome:'callcommand-dobott', error:String(err&&err.message||err) }) }
      setTimeout(() => finish({ ok:false, outcome:'callback-timeout', error:'sheet copy inventory callback timed out' }), timeout)
    })
    if (!inventory?.ok) return fail(inventory?.outcome || 'inventory-failed', inventory?.error || 'could not resolve live worksheet inventory')
    const names = inventory.names || []
    const hits = []
    for (let i=0;i<names.length;i++) if (names[i]===op.sheet) hits.push(i)
    if (hits.length !== 1) return fail('sheet-identity-ambiguous', 'source worksheet identity is not unique at dispatch boundary', { sheet:op.sheet, count:hits.length })
    if (names.includes(op.name)) return fail('already-exists', 'target worksheet name already exists', { sheet:op.name })
    const sourceIndex = hits[0]
    let ret
    try { ret = e.asc_copyWorksheet(-1, [op.name], [sourceIndex]) }
    catch (err) { return fail('operation-error', String(err&&err.message||err), { sheet:op.sheet, name:op.name, sourceIndex }) }
    if (ret === false) return fail('copy-failed', 'asc_copyWorksheet returned false', { sheet:op.sheet, name:op.name, sourceIndex })
    const deadline = Date.now()+timeout
    while (Date.now() <= deadline) {
      const count = typeof e.asc_getWorksheetsCount === 'function' ? e.asc_getWorksheetsCount() : null
      if (typeof e.asc_getWorksheetName === 'function' && Number.isInteger(count)) {
        let targetCount=0
        for(let i=0;i<count;i++) if(e.asc_getWorksheetName(i)===op.name) targetCount++
        if(targetCount===1)return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:'sheet.copy',sheet:op.sheet,name:op.name,sourceIndex,dispatchApi:'asc_copyWorksheet'}
        if(targetCount>1)return fail('verification-failed','target worksheet identity became ambiguous',{sheet:op.sheet,name:op.name,sourceIndex,targetCount})
      }
      await new Promise(r=>setTimeout(r,25))
    }
    return fail('verification-failed','asc_copyWorksheet returned without exposing the target worksheet',{sheet:op.sheet,name:op.name,sourceIndex})
  }, { u: apiHely, op: operation, timeout: timeoutMs })
}

async function runOperationInFrame(frame, apiHely, operation, timeoutMs = 15000) {
  if (operation?.type === 'sheet.create') return runNativeSheetCreateInFrame(frame, apiHely, operation, timeoutMs)
  if (operation?.type === 'sheet.move') return runNativeSheetMoveInFrame(frame, apiHely, operation, timeoutMs)
  if (operation?.type === 'sheet.copy') return runNativeSheetCopyInFrame(frame, apiHely, operation, timeoutMs)
  const body = `return (${operationCommand.toString()})(${JSON.stringify(operation)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: 'live-coedit-editor', error: 'callCommand is unavailable' })
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try {
      editor.callCommand(new Function(commandBody), false, (value) => finish(value === undefined ? { ok: false, outcome: 'ures-callback', source: 'live-coedit-editor', error: 'callCommand returned undefined' } : value))
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

module.exports = { operationCommand, runNativeSheetCreateInFrame, runNativeSheetMoveInFrame, runNativeSheetCopyInFrame, runOperationInFrame, runOperationLive }
