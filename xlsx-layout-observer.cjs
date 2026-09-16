'use strict'
function layoutObserveCommand(spec,apply){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function bad(outcome,error,extra){return Object.assign({ok:false,outcome,source:'live-coedit-editor',authority:'LIVE_READ',error},extra||{})}
  try{
    if(!spec||!spec.sheet||!spec.range||!spec.type)return bad('invalid-layout-spec','sheet, range and type are required')
    const sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh)return bad('sheet-not-found','sheet not found')
    const r=has(sh,'GetRange')?sh.GetRange(spec.range):null;if(!r)return bad('range-not-found','range not found')
    const col=spec.type==='column.width'||spec.type==='columns.hidden', row=spec.type==='row.height'||spec.type==='rows.hidden'
    const t=col&&has(r,'GetEntireColumn')?r.GetEntireColumn():row&&has(r,'GetEntireRow')?r.GetEntireRow():r
    const eq=(a,b)=>typeof b==='number'?typeof a==='number'&&Math.abs(a-b)<=0.05:a===b
    if(spec.type==='columns.autofit'||spec.type==='rows.autofit'){
      const isColumns=spec.type==='columns.autofit',getter=isColumns?'GetColumnWidth':'GetRowHeight'
      // Measured 9.3.4.60 runtime contract: AutoFit must be invoked on the
      // direct ApiWorksheet.GetRange() object with (bRows,bCols).  The derived
      // GetEntireColumn/GetEntireRow object can expose AutoFit without applying
      // the mutation.
      if(!has(r,getter)||!has(r,'AutoFit'))return bad('layout-readback-unavailable',getter+' or AutoFit unavailable')
      const before=r[getter]()
      if(!apply){if(Number.isFinite(spec.expectedDimension))return {ok:true,outcome:eq(before,spec.expectedDimension)?'layout-already-satisfied':'layout-mismatch',source:'live-coedit-editor',authority:'LIVE_READ',noOp:eq(before,spec.expectedDimension),actual:before,expected:spec.expectedDimension};return {ok:true,outcome:'layout-autofit-ready',source:'live-coedit-editor',authority:'LIVE_READ',noOp:false,actual:before}}
      r.AutoFit(!isColumns,isColumns)
      const after=r[getter]()
      if(!Number.isFinite(after)||after<=0)return bad('layout-autofit-unverifiable','AutoFit completed but resulting dimension is not measurable',{before,actual:after})
      if(eq(before,after))return bad('layout-autofit-unchanged','AutoFit completed but measured dimension did not change; without a bound post-state this is UNKNOWN',{before,actual:after})
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
