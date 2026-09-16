'use strict'
function layoutObserveCommand(spec,apply){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function bad(outcome,error){return {ok:false,outcome,source:'live-coedit-editor',authority:'LIVE_READ',error}}
  try{
    if(!spec||!spec.sheet||!spec.range||!spec.type)return bad('invalid-layout-spec','sheet, range and type are required')
    const sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh)return bad('sheet-not-found','sheet not found')
    const r=has(sh,'GetRange')?sh.GetRange(spec.range):null;if(!r)return bad('range-not-found','range not found')
    const col=spec.type==='column.width'||spec.type==='columns.hidden', row=spec.type==='row.height'||spec.type==='rows.hidden'
    const t=col&&has(r,'GetEntireColumn')?r.GetEntireColumn():row&&has(r,'GetEntireRow')?r.GetEntireRow():r
    let getter,setter,expected
    if(spec.type==='column.width'){getter='GetColumnWidth';setter='SetColumnWidth';expected=spec.width}
    else if(spec.type==='row.height'){getter='GetRowHeight';setter='SetRowHeight';expected=spec.height}
    else if(spec.type==='columns.hidden'||spec.type==='rows.hidden'){getter='GetHidden';setter='SetHidden';expected=spec.hidden}
    else return bad('unsupported-layout-type','bounded observer supports width, height and hidden state')
    if(!has(t,getter)||!has(t,setter))return bad('layout-readback-unavailable',getter+' or '+setter+' unavailable')
    const before=t[getter]();
    const eq=(a,b)=>typeof b==='number'?typeof a==='number'&&Math.abs(a-b)<=0.05:a===b
    if(!apply)return {ok:true,outcome:eq(before,expected)?'layout-already-satisfied':'layout-mismatch',source:'live-coedit-editor',authority:'LIVE_READ',noOp:eq(before,expected),before,actual:before,expected}
    if(eq(before,expected))return {ok:true,outcome:'layout-already-satisfied',source:'live-coedit-editor',authority:'LIVE_VERIFY',noOp:true,before,actual:before,expected,applied:[]}
    t[setter](expected);const after=t[getter]();const pass=eq(after,expected)
    return {ok:pass,outcome:pass?'layout-live-verified':'layout-live-verify-failed',source:'live-coedit-editor',authority:pass?'LIVE_VERIFY':'LIVE_READ',noOp:false,before,actual:after,expected,applied:pass?[spec.type]:[],verification:{status:pass?'PASS':'FAIL',expected,actual:after}}
  }catch(err){return bad('layout-observer-error',String(err&&err.message||err))}
}
module.exports={layoutObserveCommand}
