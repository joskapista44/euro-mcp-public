'use strict'

// M4.6 targeted live-runtime probe.
// This deliberately does not mutate the workbook. It only inspects whether the
// EuroOffice/ONLYOFFICE runtime exposes the conditional-formatting object model
// required for machine-verifiable M4.6 acceptance.

function conditionalFormattingCapabilityCommand(sheetName, rangeAddress) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function methods(o, names) {
    var out = {}
    for (var i = 0; i < names.length; i++) out[names[i]] = has(o, names[i])
    return out
  }
  function fail(outcome, error, extra) {
    var x = { ok: false, outcome: outcome, source: 'live-coedit-editor', error: error }
    if (extra) for (var k in extra) x[k] = extra[k]
    return x
  }

  try {
    if (!has(Api, 'GetSheet')) return fail('unsupported', 'Api.GetSheet is unavailable')
    var sheet = null
    try { sheet = Api.GetSheet(sheetName) } catch (_) {}
    if (!sheet || !has(sheet, 'GetRange')) return fail('sheet-not-found', 'worksheet/range API unavailable', { sheet: sheetName })

    var range = null
    try { range = sheet.GetRange(rangeAddress) } catch (_) {}
    if (!range) return fail('range-not-found', 'target range could not be resolved', { range: rangeAddress })

    var rangeMethods = methods(range, ['GetFormatConditions'])
    var fc = null
    var collectionError = null
    if (rangeMethods.GetFormatConditions) {
      try { fc = range.GetFormatConditions() } catch (err) { collectionError = String(err && err.message ? err.message : err) }
    }

    var collectionMethods = methods(fc, [
      'Add', 'AddAboveAverage', 'AddColorScale', 'AddDatabar',
      'AddIconSetCondition', 'AddTop10', 'AddUniqueValues',
      'Delete', 'GetCount', 'GetItem'
    ])

    var count = null
    var countError = null
    if (fc && collectionMethods.GetCount) {
      try { count = fc.GetCount() } catch (err) { countError = String(err && err.message ? err.message : err) }
    }

    var required = {
      GetFormatConditions: !!rangeMethods.GetFormatConditions,
      Add: !!collectionMethods.Add,
      GetCount: !!collectionMethods.GetCount,
      GetItem: !!collectionMethods.GetItem
    }
    var missing = []
    for (var key in required) if (!required[key]) missing.push(key)

    return {
      ok: missing.length === 0,
      outcome: missing.length === 0 ? 'capability-present' : 'unsupported',
      source: 'live-coedit-editor',
      sheet: sheetName,
      range: String(rangeAddress).toUpperCase(),
      required: required,
      missing: missing,
      rangeMethods: rangeMethods,
      collectionMethods: collectionMethods,
      existingRuleCount: count,
      collectionError: collectionError,
      countError: countError
    }
  } catch (err) {
    return fail('probe-error', String(err && err.message ? err.message : err))
  }
}

async function runCapabilityProbeInFrame(frame, apiHely, sheetName, rangeAddress, callbackTimeoutMs = 15000) {
  const body = `return (${conditionalFormattingCapabilityCommand.toString()})(${JSON.stringify(sheetName)}, ${JSON.stringify(rangeAddress)});`
  return frame.evaluate(({ apiHely, body, timeout }) => new Promise((resolve) => {
    const editor = apiHely === 'window.editor' ? window.editor : (window.Asc || {}).editor
    let done = false
    const finish = (value) => { if (!done) { done = true; resolve(value) } }
    try {
      editor.callCommand(new Function(body), false, (value) => finish(value === undefined
        ? { ok: false, outcome: 'empty-callback', source: 'live-coedit-editor', error: 'callCommand callback returned undefined' }
        : value))
      setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', source: 'live-coedit-editor', error: 'callCommand callback timed out' }), timeout)
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-error', source: 'live-coedit-editor', error: String(err && err.message ? err.message : err) })
    }
  }), { apiHely, body, timeout: callbackTimeoutMs })
}

module.exports = {
  conditionalFormattingCapabilityCommand,
  runCapabilityProbeInFrame,
}
