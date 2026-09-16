'use strict'
function layoutObserveCommand(spec,apply){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function bad(outcome,error){return {ok:false,outcome,source:'live-coedit-editor',authority:'LIVE_READ',error}}
  try{
    if(!spec||!spec.sheet||!spec.range||!spec.type)return bad('invalid-layout-spec','sheet, range and type are required')
    const sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh)return bad('sheet-not-found','sheet not found')
    const r=has(sh,'GetRange')?sh.GetRange(spec.range):null;if(!r)return bad('range-not-found','range not found')
    const col=spec.type==='column.width'||spec.type==='columns.hidden'||spec.type==='columns.autofit', row=spec.type==='row.height'||spec.type==='rows.hidden'||spec.type==='rows.autofit'
    const t=col&&has(r,'GetEntireColumn')?r.GetEntireColumn():row&&has(r,'GetEntireRow')?r.GetEntireRow():r
    const eq=(a,b)=>typeof b==='number'?typeof a==='number'&&Math.abs(a-b)<=0.05:a===b
    if(spec.type==='columns.autofit'||spec.type==='rows.autofit'){
      const getter=spec.type==='columns.autofit'?'GetColumnWidth':'GetRowHeight',setter=spec.type==='columns.autofit'?'AutoFit':'AutoFit'
      if(!has(t,getter)||!has(t,setter))return bad('layout-readback-unavailable',getter+' or '+setter+' unavailable')
      const actual=t[getter]();
      if(!apply){if(Number.isFinite(spec.expectedDimension))return {ok:true,outcome:eq(actual,spec.expectedDimension)?'layout-already-satisfied':'layout-mismatch',source:'live-coedit-editor',authority:'LIVE_READ',noOp:eq(actual,spec.expectedDimension),actual,expected:spec.expectedDimension};return {ok:true,outcome:'layout-autofit-ready',source:'live-coedit-editor',authority:'LIVE_READ',noOp:false,actual}}
      const before=actual;t[setter]();const after=t[getter]();
      if(!Number.isFinite(before)||!Number.isFinite(after)||eq(before,after))return bad('layout-autofit-unverifiable','AutoFit did not produce a measurable dimension change')
      return {ok:true,outcome:'layout-live-verified',source:'live-coedit-editor',authority:'LIVE_VERIFY',noOp:false,before,actual:after,expected:after,applied:[spec.type],verification:{status:'PASS',before,actual:after}}
    }
    let getter,setter,expected
    if(spec.type==='column.width'){getter='GetColumnWidth';setter='SetColumnWidth';expected=spec.width}
    else if(spec.type==='row.height'){getter='GetRowHeight';setter='SetRowHeight';expected=spec.height}
    else if(spec.type==='columns.hidden'||spec.type==='rows.hidden'){getter='GetHidden';setter='SetHidden';expected=spec.hidden}
    else return bad('unsupported-layout-type','bounded observer supports width, height, hidden state and AutoFit')
    if(!has(t,getter)||!has(t,setter))return bad('layout-readback-unavailable',getter+' or '+setter+' unavailable')
    const before=t[getter]();
    if(!apply)return {ok:true,outcome:eq(before,expected)?'layout-already-satisfied':'layout-mismatch',source:'live-coedit-editor',authority:'LIVE_READ',noOp:eq(before,expected),before,actual:before,expected}
    if(eq(before,expected))return {ok:true,outcome:'layout-already-satisfied',source:'live-coedit-editor',authority:'LIVE_VERIFY',noOp:true,before,actual:before,expected,applied:[]}
    t[setter](expected);const after=t[getter]();const pass=eq(after,expected)
    return {ok:pass,outcome:pass?'layout-live-verified':'layout-live-verify-failed',source:'live-coedit-editor',authority:pass?'LIVE_VERIFY':'LIVE_READ',noOp:false,before,actual:after,expected,applied:pass?[spec.type]:[],verification:{status:pass?'PASS':'FAIL',expected,actual:after}}
  }catch(err){return bad('layout-observer-error',String(err&&err.message||err))}
}
module.exports={layoutObserveCommand}
