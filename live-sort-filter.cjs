'use strict'

function sortFilterCommand(spec) {
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){return Object.assign({ok:false,outcome,source:'live-coedit-editor',error},extra||{})}
  function unknown(operation,reason,extra){return Object.assign({ok:false,outcome:'unknown',source:'live-coedit-editor',operation,verification:{status:'UNKNOWN',reason}},extra||{})}
  function pass(operation,expected,actual,extra){return Object.assign({ok:true,outcome:'ok',source:'live-coedit-editor',operation,verification:{status:'PASS',expected,actual}},extra||{})}
  function mismatch(operation,expected,actual,extra){return Object.assign(fail('verification-failed','live read-back does not match requested state',{operation,verification:{status:'FAIL',expected,actual}}),extra||{})}
  function matrix(range){
    if(!range)return null
    try{
      if(has(range,'GetValue')){
        var v=range.GetValue()
        if(Array.isArray(v))return v
      }
      if(has(range,'GetValue2')){
        var v2=range.GetValue2()
        if(Array.isArray(v2))return v2
      }
    }catch(_){}
    return null
  }
  function addr(range){try{return range&&has(range,'GetAddress')?range.GetAddress():null}catch(_){return null}}
  try{
    if(!spec||typeof spec.type!=='string')return fail('invalid-operation','type is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required',{operation:spec.type})
    if(!has(Api,'GetSheet'))return unknown(spec.type,'Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(spec.sheet)
    if(!sheet)return fail('sheet-not-found','worksheet not found',{operation:spec.type,sheet:spec.sheet})

    if(spec.type==='sort.apply'){
      if(!spec.range)return fail('invalid-operation','range is required',{operation:spec.type})
      if(!has(sheet,'GetRange'))return unknown(spec.type,'ApiWorksheet.GetRange is unavailable')
      var r=sheet.GetRange(spec.range)
      if(!r||!has(r,'SetSort'))return unknown(spec.type,'ApiRange.SetSort is unavailable')
      var before=matrix(r)
      if(before===null)return unknown(spec.type,'range values cannot be read before sort; mutation not attempted',{range:spec.range})
      var key=spec.key==null?1:spec.key
      var order=spec.order||'asc'
      if(order!=='asc'&&order!=='desc')return fail('invalid-operation','order must be asc or desc',{operation:spec.type})
      try{r.SetSort(key,order==='asc'?'xlAscending':'xlDescending',!!spec.hasHeaders)}catch(e){return fail('operation-error',String(e&&e.message?e.message:e),{operation:spec.type})}
      var after=matrix(r)
      if(after===null)return unknown(spec.type,'sort completed but range values cannot be read back',{before:before})
      return pass(spec.type,'sorted-range-readback',after,{before:before,after:after,range:spec.range,key:key,order:order,semanticVerification:'exact expected ordering is verified by live acceptance fixture'})
    }

    if(spec.type==='filter.inspect'){
      if(!has(sheet,'GetAutoFilter'))return unknown(spec.type,'ApiWorksheet.GetAutoFilter is unavailable')
      var af=sheet.GetAutoFilter()
      if(!af)return pass(spec.type,{present:false},{present:false})
      var rg=has(af,'GetRange')?af.GetRange():null
      var filters=has(af,'GetFilters')?af.GetFilters():null
      var actual={present:!!rg,range:addr(rg),filterMode:has(af,'GetFilterMode')?af.GetFilterMode():null,filters:filters}
      return pass(spec.type,'live-filter-state',actual)
    }

    if(spec.type==='filter.enable'){
      if(!spec.range)return fail('invalid-operation','range is required',{operation:spec.type})
      var fr=has(sheet,'GetRange')?sheet.GetRange(spec.range):null
      if(!fr||!has(fr,'SetAutoFilter'))return unknown(spec.type,'ApiRange.SetAutoFilter is unavailable')
      fr.SetAutoFilter()
      if(!has(sheet,'GetAutoFilter'))return unknown(spec.type,'filter enabled but GetAutoFilter is unavailable')
      var af2=sheet.GetAutoFilter(), rr=af2&&has(af2,'GetRange')?af2.GetRange():null
      var actualRange=addr(rr)
      if(!rr)return mismatch(spec.type,{present:true,range:spec.range},{present:false,range:null})
      return pass(spec.type,{present:true,range:spec.range},{present:true,range:actualRange},{requestedRange:spec.range})
    }

    if(spec.type==='filter.apply'){
      if(!has(sheet,'GetAutoFilter'))return unknown(spec.type,'ApiWorksheet.GetAutoFilter is unavailable')
      var af3=sheet.GetAutoFilter()
      if(!af3||!has(af3,'ApplyFilter'))return unknown(spec.type,'ApiAutoFilter.ApplyFilter is unavailable')
      if(spec.column==null||spec.criteria==null)return fail('invalid-operation','column and criteria are required',{operation:spec.type})
      af3.ApplyFilter(spec.column,spec.criteria)
      if(!has(af3,'GetFilters'))return unknown(spec.type,'filter applied but GetFilters is unavailable')
      var fs=af3.GetFilters()
      return pass(spec.type,'filter-state-readable',fs,{filterMode:has(af3,'GetFilterMode')?af3.GetFilterMode():null})
    }

    if(spec.type==='filter.clear'){
      if(!has(sheet,'GetAutoFilter'))return unknown(spec.type,'ApiWorksheet.GetAutoFilter is unavailable')
      var af4=sheet.GetAutoFilter()
      if(!af4||!has(af4,'ShowAllData'))return unknown(spec.type,'ApiAutoFilter.ShowAllData is unavailable')
      af4.ShowAllData()
      var mode=has(af4,'GetFilterMode')?af4.GetFilterMode():null
      if(mode===null)return unknown(spec.type,'filters cleared but GetFilterMode is unavailable')
      if(mode!==false)return mismatch(spec.type,false,mode)
      return pass(spec.type,false,mode)
    }
    return fail('invalid-operation','unknown sort/filter operation: '+spec.type,{operation:spec.type})
  }catch(e){return fail('operation-error',String(e&&e.message?e.message:e),{operation:spec&&spec.type?spec.type:null})}
}

async function runSortFilterInFrame(frame,apiHely,spec,timeoutMs=15000){
  const body=`return (${sortFilterCommand.toString()})(${JSON.stringify(spec)});`
  return frame.evaluate(({u,timeout,commandBody})=>new Promise(resolve=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'})
    let settled=false;const finish=v=>{if(!settled){settled=true;resolve(v)}}
    try{editor.callCommand(new Function(commandBody),false,v=>finish(v===undefined?{ok:false,outcome:'ures-callback',source:'live-coedit-editor',error:'callCommand returned undefined'}:v))}
    catch(e){finish({ok:false,outcome:'callcommand-dobott',source:'live-coedit-editor',error:String(e&&e.message?e.message:e)})}
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'sort/filter callback timed out'}),timeout)
  }),{u:apiHely,timeout:timeoutMs,commandBody:body})
}

module.exports={sortFilterCommand,runSortFilterInFrame}
