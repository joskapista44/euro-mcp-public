'use strict'

// M1.1 Workbook Inspector -- READS workbook structure from the LIVE spreadsheet editor.
// It deliberately does not download/unzip the xlsx and does not invoke DocBuilder: an already
// existing document belongs to the co-editing route, so the source of truth here is the editor's
// current Api object (including changes that may not have been flushed to WebDAV yet).
//
// The editor object is resolved by the caller exactly as coedit.cjs already does for writes.
// `callCommand` is used with a callback because ONLYOFFICE runs the Office API function in an
// isolated context; only plain JSON-shaped data may cross back to us.
async function inspectWorkbookInFrame(frame, apiHely, timeoutMs = 10000) {
  return frame.evaluate(({ u, timeout }) => new Promise((resolve) => {
    const editor = u === 'window.editor' ? window.editor : (window.Asc || {}).editor
    if (!editor || typeof editor.callCommand !== 'function') {
      resolve({ ok: false, outcome: 'nincs-api', indok: 'callCommand nem erheto el a megadott editor objektumon' })
      return
    }

    let settled = false
    const finish = (value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(value)
    }
    const timer = setTimeout(() => finish({ ok: false, outcome: 'callback-timeout', indok: 'a workbook inspector callCommand callback nem hivodott meg idoben' }), timeout)

    try {
      editor.callCommand(function () {
        try {
          function safeCall(obj, name, fallback) {
            try {
              return obj && typeof obj[name] === 'function' ? obj[name]() : fallback
            } catch (_) { return fallback }
          }
          function rangeAddress(range) {
            if (!range) return null
            try { return typeof range.GetAddress === 'function' ? range.GetAddress() : null } catch (_) { return null }
          }
          function rangeDimensions(range) {
            if (!range) return { rows: null, columns: null }
            var rows = null
            var columns = null
            try {
              var rr = typeof range.GetRows === 'function' ? range.GetRows() : null
              if (rr && typeof rr.GetCount === 'function') rows = rr.GetCount()
            } catch (_) {}
            try {
              var cc = typeof range.GetCols === 'function' ? range.GetCols() : null
              if (cc && typeof cc.GetCount === 'function') columns = cc.GetCount()
            } catch (_) {}
            return { rows: rows, columns: columns }
          }
          function nameInfo(n) {
            var name = null
            var ref = null
            try { if (n && typeof n.GetName === 'function') name = n.GetName() } catch (_) {}
            try {
              if (n && typeof n.GetRefersTo === 'function') ref = n.GetRefersTo()
              else if (n && typeof n.GetRefersToRange === 'function') ref = rangeAddress(n.GetRefersToRange())
            } catch (_) {}
            return { name: name, ref: ref }
          }

          var sheets = Api.GetSheets()
          var active = Api.GetActiveSheet()
          var activeName = active && typeof active.GetName === 'function' ? active.GetName() : null
          var resultSheets = []

          for (var i = 0; i < sheets.length; i++) {
            var s = sheets[i]
            var used = safeCall(s, 'GetUsedRange', null)
            var dims = rangeDimensions(used)
            var tablesRaw = safeCall(s, 'GetListObjects', []) || []
            var tables = []
            for (var t = 0; t < tablesRaw.length; t++) {
              var table = tablesRaw[t]
              var tableName = null
              var tableRange = null
              try {
                if (table && typeof table.GetDisplayName === 'function') tableName = table.GetDisplayName()
                else if (table && typeof table.GetName === 'function') tableName = table.GetName()
              } catch (_) {}
              try {
                if (table && typeof table.GetRange === 'function') tableRange = rangeAddress(table.GetRange())
              } catch (_) {}
              tables.push({ name: tableName, range: tableRange })
            }

            var af = safeCall(s, 'GetAutoFilter', null)
            var autoFilter = { present: false, range: null, filterMode: null }
            if (af) {
              autoFilter.present = true
              try { if (typeof af.GetRange === 'function') autoFilter.range = rangeAddress(af.GetRange()) } catch (_) {}
              try { if (typeof af.GetFilterMode === 'function') autoFilter.filterMode = af.GetFilterMode() } catch (_) {}
            }

            var fp = safeCall(s, 'GetFreezePanes', null)
            var freezePanes = { present: !!fp, type: null }
            try {
              if (typeof Api.GetFreezePanesType === 'function') freezePanes.type = Api.GetFreezePanesType()
            } catch (_) {}

            var localNamesRaw = safeCall(s, 'GetDefNames', []) || []
            var localNames = []
            for (var n = 0; n < localNamesRaw.length; n++) localNames.push(nameInfo(localNamesRaw[n]))

            var sheetName = safeCall(s, 'GetName', null)
            resultSheets.push({
              index: i,
              name: sheetName,
              visible: safeCall(s, 'GetVisible', null),
              active: sheetName === activeName,
              usedRange: { address: rangeAddress(used), rows: dims.rows, columns: dims.columns },
              tables: tables,
              autoFilter: autoFilter,
              freezePanes: freezePanes,
              definedNames: localNames,
            })
          }

          var workbookNamesRaw = typeof Api.GetDefNames === 'function' ? (Api.GetDefNames() || []) : []
          var workbookNames = []
          for (var w = 0; w < workbookNamesRaw.length; w++) workbookNames.push(nameInfo(workbookNamesRaw[w]))

          return {
            ok: true,
            outcome: 'ok',
            source: 'live-coedit-editor',
            activeSheet: activeName,
            sheetCount: resultSheets.length,
            sheets: resultSheets,
            definedNames: workbookNames,
          }
        } catch (err) {
          return { ok: false, outcome: 'inspector-hiba', indok: String(err && err.message ? err.message : err) }
        }
      }, function (value) {
        if (!value || typeof value !== 'object') {
          finish({ ok: false, outcome: 'ures-callback', indok: 'a workbook inspector nem kapott strukturalt visszateresi erteket' })
          return
        }
        finish(value)
      })
    } catch (err) {
      finish({ ok: false, outcome: 'callcommand-dobott', indok: String(err && err.message ? err.message : err) })
    }
  }), { u: apiHely, timeout: timeoutMs })
}

module.exports = { inspectWorkbookInFrame }
