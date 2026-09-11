'use strict';

function pivotCommand(spec) {
  function safe(fn) {
    try { return { ok: true, value: fn() }; }
    catch (e) { return { ok: false, error: String(e && e.message ? e.message : e) }; }
  }
  function call(obj, name) {
    var r = safe(function () { return obj && typeof obj[name] === 'function' ? obj[name]() : null; });
    return r.ok ? r.value : null;
  }
  function sourceAddress(p) {
    var src = call(p, 'GetSource');
    if (src && typeof src.GetAddress === 'function') {
      var a = safe(function () { return src.GetAddress(); });
      return a.ok ? a.value : null;
    }
    return typeof src === 'string' ? src : null;
  }
  function parentSheetName(p) {
    var parent = call(p, 'GetParent');
    if (parent && typeof parent.GetName === 'function') {
      var n = safe(function () { return parent.GetName(); });
      return n.ok ? n.value : null;
    }
    return null;
  }
  function snap(p) {
    if (!p) return null;
    return {
      name: call(p, 'GetName'), source: sourceAddress(p), parentSheet: parentSheetName(p),
      rowFields: (call(p, 'GetRowFields') || []).length,
      columnFields: (call(p, 'GetColumnFields') || []).length,
      dataFields: (call(p, 'GetDataFields') || []).length,
      styleName: call(p, 'GetStyleName'), title: call(p, 'GetTitle'), description: call(p, 'GetDescription')
    };
  }
  function get(name) {
    try { return typeof Api.GetPivotByName === 'function' ? Api.GetPivotByName(name) : null; }
    catch (_) { return null; }
  }
  function scalar(v) { return v === null || typeof v === 'undefined' ? null : String(v); }
  function getData(p, items) {
    var r=safe(function(){ return p.GetData(items); });
    return r.ok ? {ok:true,value:scalar(r.value)} : {ok:false,error:r.error};
  }

  try {
    var op = spec.operation, p, before, after;
    if (op === 'pivot.inspect') {
      p = get(spec.name);
      return { ok: true, outcome: 'ok', source: 'live-coedit-editor', operation: op,
        verification: { status: 'PASS', expected: spec.expectedPresent === true ? 'pivot-present' : 'live-pivot-state', actual: snap(p) } };
    }
    if (op === 'pivot.createNewWorksheet') {
      before = (Api.GetAllPivotTables() || []).length;
      var src = Api.GetActiveSheet().GetRange(spec.source);
      p = Api.InsertPivotNewWorksheet(src);
      if (!p) return { ok: false, outcome: 'unsupported', source: 'live-coedit-editor', operation: op,
        verification: { status: 'UNKNOWN', expected: { count: before + 1 }, actual: null } };
      if (spec.name && typeof p.SetName === 'function') p.SetName(spec.name);
      after = (Api.GetAllPivotTables() || []).length;
      var a = snap(p);
      var pass = after === before + 1 && a && (!spec.name || a.name === spec.name) && a.source === spec.source;
      return { ok: pass, outcome: pass ? 'ok' : 'verification-failed', source: 'live-coedit-editor', operation: op,
        verification: { status: pass ? 'PASS' : 'FAIL', expected: { count: before + 1, name: spec.name || null, source: spec.source }, actual: { count: after, pivot: a } } };
    }
    p = get(spec.name);
    if (!p) return { ok: false, outcome: 'not-found', source: 'live-coedit-editor', operation: op,
      verification: { status: 'FAIL', expected: spec.name, actual: null } };
    if (op === 'pivot.addFields') {
      before = snap(p); var fieldSpec = {};
      if (spec.rowFields && spec.rowFields.length) fieldSpec.rows = spec.rowFields.length === 1 ? spec.rowFields[0] : spec.rowFields;
      if (spec.columnFields && spec.columnFields.length) fieldSpec.columns = spec.columnFields.length === 1 ? spec.columnFields[0] : spec.columnFields;
      if (spec.pageFields && spec.pageFields.length) fieldSpec.pages = spec.pageFields.length === 1 ? spec.pageFields[0] : spec.pageFields;
      p.AddFields(fieldSpec); after = snap(p);
      var er = (spec.rowFields || []).length, ec = (spec.columnFields || []).length;
      var pass2 = after.rowFields === before.rowFields + er && after.columnFields === before.columnFields + ec;
      return { ok: pass2, outcome: pass2 ? 'ok' : 'verification-failed', source: 'live-coedit-editor', operation: op,
        verification: { status: pass2 ? 'PASS' : 'FAIL', expected: { rowFields: before.rowFields + er, columnFields: before.columnFields + ec }, actual: after } };
    }
    if (op === 'pivot.addDataField') {
      before = snap(p); p.AddDataField(spec.field); after = snap(p);
      var pass3 = after.dataFields === before.dataFields + 1;
      return { ok: pass3, outcome: pass3 ? 'ok' : 'verification-failed', source: 'live-coedit-editor', operation: op,
        verification: { status: pass3 ? 'PASS' : 'FAIL', expected: { dataFields: before.dataFields + 1 }, actual: after } };
    }
    if (op === 'pivot.rename') {
      before = snap(p); p.SetName(spec.newName); after = snap(get(spec.newName) || p);
      var pass4 = !!after && after.name === spec.newName && after.source === before.source;
      return { ok: pass4, outcome: pass4 ? 'ok' : 'verification-failed', source: 'live-coedit-editor', operation: op,
        verification: { status: pass4 ? 'PASS' : 'FAIL', expected: { name: spec.newName, source: before.source }, actual: after, before: before } };
    }
    if (op === 'pivot.style') {
      p.SetStyleName(spec.styleName); after = snap(p); var pass5 = after.styleName === spec.styleName;
      return { ok: pass5, outcome: pass5 ? 'ok' : 'verification-failed', source: 'live-coedit-editor', operation: op,
        verification: { status: pass5 ? 'PASS' : 'FAIL', expected: { styleName: spec.styleName }, actual: after } };
    }
    if (op === 'pivot.refresh') {
      var assertions=Array.isArray(spec.assertions)?spec.assertions:[];
      if(!assertions.length){
        p.RefreshTable(); after=snap(p);
        return {ok:false,outcome:'unverified',source:'live-coedit-editor',operation:op,
          verification:{status:'UNKNOWN',expected:'one or more public GetData semantic assertions',actual:after}};
      }
      var beforeData=assertions.map(function(x){return {items:x.items,readback:getData(p,x.items)};});
      if(spec.sourceMutation&&spec.sourceMutation.range){
        var parent=Api.GetSheet(spec.sourceMutation.sheetName)||Api.GetActiveSheet();
        parent.GetRange(spec.sourceMutation.range).SetValue(spec.sourceMutation.value);
      }
      p.RefreshTable();
      var afterData=assertions.map(function(x){return {items:x.items,readback:getData(p,x.items),expected:scalar(x.expectedAfter),expectedBefore:typeof x.expectedBefore==='undefined'?undefined:scalar(x.expectedBefore)};});
      var checks=afterData.map(function(x,i){
        var beforeValue=beforeData[i].readback&&beforeData[i].readback.value;
        var afterValue=x.readback&&x.readback.value;
        var beforeOk=typeof x.expectedBefore==='undefined'||beforeValue===x.expectedBefore;
        var afterOk=!!x.readback.ok&&afterValue===x.expected;
        return {items:x.items,before:beforeValue,after:afterValue,expectedBefore:x.expectedBefore,expectedAfter:x.expected,beforeOk:beforeOk,afterOk:afterOk,pass:beforeOk&&afterOk};
      });
      var refreshPass=checks.length>0&&checks.every(function(x){return x.pass;});
      return {ok:refreshPass,outcome:refreshPass?'ok':'verification-failed',source:'live-coedit-editor',operation:op,
        verification:{status:refreshPass?'PASS':'FAIL',expected:{semanticGetDataAssertions:assertions.length},actual:{checks:checks,pivot:snap(p)},readback:'public-ApiPivotTable.GetData'}};
    }
    return { ok: false, outcome: 'unknown-operation', source: 'live-coedit-editor', operation: op,
      verification: { status: 'FAIL', expected: 'known pivot operation', actual: op } };
  } catch (e) {
    return { ok: false, outcome: 'execution-error', source: 'live-coedit-editor', operation: spec && spec.operation,
      verification: { status: 'UNKNOWN', expected: 'operation completes with semantic readback', actual: String(e && e.message ? e.message : e) } };
  }
}

async function runPivotInFrame(frame, editorExpr, spec) {
  var body = 'return (' + pivotCommand.toString() + ')(' + JSON.stringify(spec) + ');';
  return frame.evaluate(({ editorExpr, body }) => new Promise(resolve => {
    var editor;
    try { editor = Function('return (' + editorExpr + ')')(); }
    catch (e) { return resolve({ ok: false, outcome: 'editor-unavailable', source: 'live-coedit-editor', verification: { status: 'UNKNOWN', actual: String(e) } }); }
    if (!editor || typeof editor.callCommand !== 'function') return resolve({ ok: false, outcome: 'editor-unavailable', source: 'live-coedit-editor', verification: { status: 'UNKNOWN' } });
    editor.callCommand(new Function(body), false, resolve);
  }), { editorExpr, body });
}

module.exports = { pivotCommand, runPivotInFrame };
