'use strict'

const { runInLiveEditor } = require('./live-formatting.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function validateLayoutRequest(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return { ok:false, outcome:'invalid-layout', error:'layout request must be an object' }
  const allowed = new Set(['columnWidth','rowHeight','autofitRows','autofitColumns','hidden','structure'])
  for (const key of Object.keys(input)) if (!allowed.has(key)) return { ok:false, outcome:'invalid-layout', error:`unknown layout property: ${key}` }
  if (input.columnWidth !== undefined && (typeof input.columnWidth !== 'number' || !Number.isFinite(input.columnWidth) || input.columnWidth <= 0)) return { ok:false, outcome:'invalid-layout', error:'columnWidth must be a positive finite number' }
  if (input.rowHeight !== undefined && (typeof input.rowHeight !== 'number' || !Number.isFinite(input.rowHeight) || input.rowHeight <= 0)) return { ok:false, outcome:'invalid-layout', error:'rowHeight must be a positive finite number' }
  for (const key of ['autofitRows','autofitColumns','hidden']) if (input[key] !== undefined && typeof input[key] !== 'boolean') return { ok:false, outcome:'invalid-layout', error:`${key} must be boolean` }
  if (input.structure !== undefined && !['rows.insert','rows.delete','columns.insert','columns.delete'].includes(input.structure)) return { ok:false, outcome:'invalid-layout', error:'structure must be rows.insert, rows.delete, columns.insert or columns.delete' }
  if (!Object.keys(input).length) return { ok:false, outcome:'invalid-layout', error:'layout request must contain at least one operation' }
  return { ok:true }
}

function rowColumnCommand(sheetName, rangeAddress, options) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra)for(var k in extra)x[k]=extra[k]; return x }
  function read(range, method, prop){
    try { if (method && has(range,method)) return { measurable:true, value:range[method]() }; if (prop && range[prop] !== undefined) return { measurable:true, value:range[prop] } } catch (_) {}
    return { measurable:false, value:null }
  }
  try {
    if(!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=null; try{ sheet=Api.GetSheet(sheetName) }catch(_){}
    if(!sheet || !has(sheet,'GetRange')) return fail('sheet-not-found','worksheet/range API unavailable',{sheet:sheetName})
    var range=null; try{ range=sheet.GetRange(rangeAddress) }catch(_){}
    if(!range) return fail('range-not-found','target range could not be resolved',{range:rangeAddress})

    var applied=[]; var checks={}; var mismatches=[]; var unknown=[]
    function verify(key, expected, getter, prop, tolerance){
      var got=read(range,getter,prop)
      if(!got.measurable){ checks[key]={status:'unknown',expected:expected,actual:null,reason:'live getter/property unavailable'}; unknown.push(key); return }
      var match = tolerance ? Math.abs(Number(got.value)-Number(expected)) <= tolerance : got.value === expected
      checks[key]={status:match?'pass':'fail',expected:expected,actual:got.value}
      if(!match)mismatches.push(key)
    }

    if(options.columnWidth!==undefined){
      if(!has(range,'SetColumnWidth')) return fail('unsupported','ApiRange.SetColumnWidth is unavailable',{capability:'columnWidth'})
      range.SetColumnWidth(options.columnWidth); applied.push('columnWidth'); verify('columnWidth',options.columnWidth,'GetColumnWidth','ColumnWidth',0.01)
    }
    if(options.rowHeight!==undefined){
      if(!has(range,'SetRowHeight')) return fail('unsupported','ApiRange.SetRowHeight is unavailable',{capability:'rowHeight'})
      var rh=range.SetRowHeight(options.rowHeight); if(rh===false)return fail('layout-error','ApiRange.SetRowHeight returned false',{capability:'rowHeight'})
      applied.push('rowHeight'); verify('rowHeight',options.rowHeight,'GetRowHeight','RowHeight',0.01)
    }
    if(options.hidden!==undefined){
      if(!has(range,'SetHidden')) return fail('unsupported','ApiRange.SetHidden is unavailable',{capability:'hidden'})
      var hiddenResult=range.SetHidden(options.hidden); if(hiddenResult===false)return fail('layout-error','ApiRange.SetHidden returned false',{capability:'hidden'})
      applied.push('hidden'); verify('hidden',options.hidden,'GetHidden','Hidden',0)
    }
    if(options.autofitRows || options.autofitColumns){
      if(!has(range,'AutoFit')) return fail('unsupported','ApiRange.AutoFit is unavailable',{capability:'autofit'})
      range.AutoFit(!!options.autofitRows,!!options.autofitColumns); applied.push('autofit')
      checks.autofit={status:'unknown',expected:{rows:!!options.autofitRows,columns:!!options.autofitColumns},actual:null,reason:'AutoFit has no direct success getter; resulting dimensions require content-dependent acceptance'}
      unknown.push('autofit')
    }
    if(options.structure){
      var isInsert=/\.insert$/.test(options.structure); var isRows=/^rows\./.test(options.structure)
      var method=isInsert?'Insert':'Delete'; var shift=isRows?(isInsert?'down':'up'):(isInsert?'right':'left')
      if(!has(range,method)) return fail('unsupported','ApiRange.'+method+' is unavailable',{capability:options.structure})
      var sr=range[method](shift); if(sr===false)return fail('layout-error','ApiRange.'+method+' returned false',{capability:options.structure})
      applied.push(options.structure)
      checks[options.structure]={status:'unknown',expected:{shift:shift},actual:null,reason:'structural mutation requires before/after range-content verification'}
      unknown.push(options.structure)
    }

    var verificationOutcome=mismatches.length?'fail':unknown.length?'unknown':'pass'
    return {ok:mismatches.length===0,outcome:mismatches.length?'verification-mismatch':'ok',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),applied:applied,verification:{outcome:verificationOutcome,checks:checks,mismatches:mismatches,unknown:unknown}}
  }catch(err){ return fail('layout-error',String(err&&err.message?err.message:err)) }
}

function rowColumnLayoutLive(common) {
  const v=validateLayoutRequest(common.layout)
  if(!v.ok) return Promise.resolve({...v,source:LIVE_SOURCE})
  return runInLiveEditor({...common,command:rowColumnCommand,args:[common.sheet,common.range,common.layout]})
}

module.exports={LIVE_SOURCE,validateLayoutRequest,rowColumnCommand,rowColumnLayoutLive}
