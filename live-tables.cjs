'use strict'

const LIVE_SOURCE = 'live-coedit-editor'

function tableCommand(spec) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function fail(outcome, error, extra) { return Object.assign({ ok: false, outcome, source: 'live-coedit-editor', error }, extra || {}) }
  function unknown(extra) { return Object.assign({ status: 'UNKNOWN', reason: 'required live table postcondition getter is unavailable' }, extra || {}) }
  function normAddress(value) {
    var s = String(value || '').replace(/\$/g, '').toUpperCase()
    var bang = s.lastIndexOf('!')
    return bang >= 0 ? s.slice(bang + 1) : s
  }
  function getSheet(name) {
    if (!has(Api, 'GetSheet')) return null
    try { return Api.GetSheet(name) } catch (_) { return null }
  }
  function tableName(table) {
    if (!table) return null
    if (has(table, 'GetName')) { try { return String(table.GetName()) } catch (_) {} }
    if (has(table, 'GetDisplayName')) { try { return String(table.GetDisplayName()) } catch (_) {} }
    return null
  }
  function tableRange(table) {
    if (!table || !has(table, 'GetRange')) return null
    try {
      var r = table.GetRange()
      if (r && has(r, 'GetAddress')) return normAddress(r.GetAddress())
    } catch (_) {}
    return null
  }
  function listTables(sheet) {
    if (!has(sheet, 'GetListObjects')) return null
    var list
    try { list = sheet.GetListObjects() } catch (_) { return null }
    if (!list || typeof list.length !== 'number') return null
    var out = []
    for (var i = 0; i < list.length; i++) out.push({ name: tableName(list[i]), range: tableRange(list[i]) })
    return out
  }

  try {
    if (!spec || typeof spec.type !== 'string') return fail('invalid-operation', 'spec.type is required')
    if (!spec.sheet) return fail('invalid-operation', 'sheet is required')
    var sheet = getSheet(spec.sheet)
    if (!sheet) return fail('sheet-not-found', 'the requested worksheet was not found', { sheet: spec.sheet })

    if (spec.type === 'table.inspect') {
      var tables = listTables(sheet)
      if (tables == null) return fail('unsupported', 'ApiWorksheet.GetListObjects is unavailable', { sheet: spec.sheet })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: spec.type, sheet: spec.sheet, tables, verification: { status: 'PASS', actual: tables } }
    }

    if (spec.type === 'table.create') {
      if (!spec.range) return fail('invalid-operation', 'table.create requires range')
      var before = listTables(sheet)
      var created = null
      var writer = null
      if (has(sheet, 'AddListObject')) {
        try { created = sheet.AddListObject(spec.range, true) } catch (err) { return fail('operation-error', String(err && err.message ? err.message : err), { writer: 'AddListObject' }) }
        writer = 'AddListObject'
      } else if (has(sheet, 'FormatAsTable')) {
        try {
          var formatted = sheet.FormatAsTable(spec.range)
          if (formatted === false) return fail('operation-error', 'ApiWorksheet.FormatAsTable returned false', { writer: 'FormatAsTable' })
        } catch (err) { return fail('operation-error', String(err && err.message ? err.message : err), { writer: 'FormatAsTable' }) }
        writer = 'FormatAsTable'
      } else return fail('unsupported', 'neither ApiWorksheet.AddListObject nor FormatAsTable is available')

      if (created && spec.name) {
        try {
          if (has(created, 'SetDisplayName')) created.SetDisplayName(spec.name)
          else if (has(created, 'SetName')) created.SetName(spec.name)
        } catch (err) { return fail('operation-error', 'table created but naming failed: ' + String(err && err.message ? err.message : err), { writer: writer }) }
      }

      var after = listTables(sheet)
      if (after == null) return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: spec.type, sheet: spec.sheet, range: spec.range, writer, verification: unknown({ expected: normAddress(spec.range) }) }
      var expectedRange = normAddress(spec.range)
      var matches = after.filter(function (t) { return t.range === expectedRange && (!spec.name || t.name === spec.name) })
      if (!matches.length) return fail('verification-failed', 'live table inventory does not contain the requested table', { writer, before, after, expected: { range: expectedRange, name: spec.name || null } })
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: spec.type, sheet: spec.sheet, range: spec.range, name: spec.name || matches[0].name, writer, before, after, verification: { status: 'PASS', expected: { range: expectedRange, name: spec.name || null }, actual: matches[0] } }
    }

    return fail('invalid-operation', 'unknown table operation: ' + spec.type)
  } catch (err) {
    return fail('table-error', String(err && err.message ? err.message : err), { operation: spec && spec.type ? spec.type : null })
  }
}

async function runTableInFrame(frame, apiHely, spec, timeoutMs = 15000) {
  const body = `return (${tableCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({ u, timeout, commandBody }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'nincs-api', source: LIVE_SOURCE, error: 'callCommand is unavailable' })
    let settled = false
    const finish = (v) => { if (!settled) { settled = true; resolve(v) } }
    try { editor.callCommand(new Function(commandBody), false, (value) => finish(value === undefined ? { ok: false, outcome: 'empty-callback', source: LIVE_SOURCE, error: 'callCommand callback returned undefined' } : value)) }
    catch (err) { finish({ ok: false, outcome: 'callcommand-error', source: LIVE_SOURCE, error: String(err && err.message ? err.message : err) }) }
    setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: LIVE_SOURCE, error: 'table callCommand callback timed out' }), timeout)
  }), { u: apiHely, timeout: timeoutMs, commandBody: body })
}

module.exports = { tableCommand, runTableInFrame }
