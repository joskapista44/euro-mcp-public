'use strict'

function sortFilterCommand(spec) {
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){return Object.assign({ok:false,outcome,source:'live-coedit-editor',error},extra||{})}
  function unknown(operation,reason,extra){return Object.assign({ok:false,outcome:'unknown',source:'live-coedit-editor',operation,verification:{status:'UNKNOWN',reason}},extra||{})}
  function pass(operation,expected,actual,extra){return Object.assign({ok:true,outcome:'ok',source:'live-coedit-editor',operation,verification:{status:'PASS',expected,actual}},extra||{})}
  function mismatch(operation,expected,actual,extra){return Object.assign(fail('verification-failed','live read-back does not match requested state',{operation,verification:{status:'FAIL',expected,actual}}),extra||{})}
  function matrix(range){try{var v=has(range,'GetValue')?range.GetValue():null;return Array.isArray(v)?v:null}catch(_){return null}}
  function addr(range){try{return range&&has(range,'GetAddress')?range.GetAddress():null}catch(_){return null}}
  function filterState(af){
    var filters=af&&has(af,'GetFilters')?af.GetFilters():null,out=[]
    if(Array.isArray(filters))for(var i=0;i<filters.length;i++){
      var f=filters[i],x={}
      try{x.operator=has(f,'GetOperator')?f.GetOperator():f.Operator}catch(_){}
      try{x.criteria1=has(f,'GetCriteria1')?f.GetCriteria1():f.Criteria1}catch(_){}
      try{x.criteria2=has(f,'GetCriteria2')?f.GetCriteria2():f.Criteria2}catch(_){}
      try{x.on=has(f,'GetOn')?f.GetOn():f.On}catch(_){}
      out.push(x)
    }
    return {filterMode:af&&has(af,'GetFilterMode')?af.GetFilterMode():null,filters:out}
  }
  try{
    if(!spec||typeof spec.type!=='string')return fail('invalid-operation','type is required')
    if(!spec.sheet)return fail('invalid-operation','sheet is required',{operation:spec.type})
    if(!has(Api,'GetSheet'))return unknown(spec.type,'Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(spec.sheet)
    if(!sheet)return fail('sheet-not-found','worksheet not found',{operation:spec.type,sheet:spec.sheet})

    if(spec.type==='sort.apply'){
      if(!spec.range||!spec.keyRange)return fail('invalid-operation','range and keyRange are required',{operation:spec.type})
      var r=has(sheet,'GetRange')?sheet.GetRange(spec.range):null
      if(!r||!has(r,'SetSort'))return unknown(spec.type,'ApiRange.SetSort is unavailable')
      var before=matrix(r);if(before===null)return unknown(spec.type,'range values cannot be read before sort; mutation not attempted')
      var order=spec.order||'asc';if(order!=='asc'&&order!=='desc')return fail('invalid-operation','order must be asc or desc',{operation:spec.type})
      var header=spec.hasHeaders===false?'xlNo':'xlYes'
      try{r.SetSort(spec.keyRange,order==='asc'?'xlAscending':'xlDescending',null,null,null,null,header,'xlSortColumns')}catch(e){return fail('operation-error',String(e&&e.message?e.message:e),{operation:spec.type})}
      var after=matrix(r);if(after===null)return unknown(spec.type,'sort completed but range values cannot be read back',{before})
      return pass(spec.type,'sorted-range-readback',after,{before,after,range:spec.range,keyRange:spec.keyRange,order})
    }

    if(spec.type==='filter.inspect'){
      if(!has(sheet,'GetAutoFilter'))return unknown(spec.type,'ApiWorksheet.GetAutoFilter is unavailable')
      var af=sheet.GetAutoFilter(),rg=af&&has(af,'GetRange')?af.GetRange():null
      var st=filterState(af);st.present=!!rg;st.range=addr(rg)
      return pass(spec.type,'live-filter-state',st)
    }

    if(spec.type==='filter.enable'){
      if(!spec.range)return fail('invalid-operation','range is required',{operation:spec.type})
      var fr=has(sheet,'GetRange')?sheet.GetRange(spec.range):null
      if(!fr||!has(fr,'SetAutoFilter'))return unknown(spec.type,'ApiRange.SetAutoFilter is unavailable')
      fr.SetAutoFilter()
      var af2=has(sheet,'GetAutoFilter')?sheet.GetAutoFilter():null,rr=af2&&has(af2,'GetRange')?af2.GetRange():null
      if(!rr)return mismatch(spec.type,{present:true},{present:false})
      return pass(spec.type,{present:true,range:spec.range},{present:true,range:addr(rr)})
    }

    if(spec.type==='filter.set'){
      if(!spec.range||spec.field==null||spec.criteria1==null)return fail('invalid-operation','range, field and criteria1 are required',{operation:spec.type})
      var tr=has(sheet,'GetRange')?sheet.GetRange(spec.range):null
      if(!tr||!has(tr,'SetAutoFilter'))return unknown(spec.type,'ApiRange.SetAutoFilter is unavailable')
      tr.SetAutoFilter(spec.field,spec.criteria1,spec.operator||'xlOr',spec.criteria2==null?undefined:spec.criteria2)
      var af3=has(sheet,'GetAutoFilter')?sheet.GetAutoFilter():null,st3=filterState(af3)
      if(!st3.filterMode||!st3.filters.length)return mismatch(spec.type,'active-filter-with-readable-criteria',st3)
      return pass(spec.type,'active-filter-with-readable-criteria',st3)
    }

    if(spec.type==='filter.reapply'){
      var af4=has(sheet,'GetAutoFilter')?sheet.GetAutoFilter():null
      if(!af4||!has(af4,'ApplyFilter'))return unknown(spec.type,'ApiAutoFilter.ApplyFilter is unavailable')
      var before4=filterState(af4);af4.ApplyFilter();var after4=filterState(af4)
      return pass(spec.type,before4,after4)
    }

    if(spec.type==='filter.clear'){
      var af5=has(sheet,'GetAutoFilter')?sheet.GetAutoFilter():null
      if(!af5||!has(af5,'ShowAllData'))return unknown(spec.type,'ApiAutoFilter.ShowAllData is unavailable')
      af5.ShowAllData();var st5=filterState(af5)
      if(st5.filters.length!==0)return mismatch(spec.type,{filters:[]},st5)
      return pass(spec.type,{filters:[]},st5)
    }
    return fail('invalid-operation','unknown sort/filter operation: '+spec.type,{operation:spec.type})
  }catch(e){return fail('operation-error',String(e&&e.message?e.message:e),{operation:spec&&spec.type?spec.type:null})}
}

async function runSortFilterInFrame(frame,apiHely,spec,timeoutMs=15000){const body=`return (${sortFilterCommand.toString()})(${JSON.stringify(spec)});`;return frame.evaluate(({u,timeout,commandBody})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'});let settled=false;const finish=v=>{if(!settled){settled=true;resolve(v)}};try{editor.callCommand(new Function(commandBody),false,v=>finish(v===undefined?{ok:false,outcome:'ures-callback',source:'live-coedit-editor',error:'callCommand returned undefined'}:v))}catch(e){finish({ok:false,outcome:'callcommand-dobott',source:'live-coedit-editor',error:String(e&&e.message?e.message:e)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'sort/filter callback timed out'}),timeout)}),{u:apiHely,timeout:timeoutMs,commandBody:body})}
module.exports={sortFilterCommand,runSortFilterInFrame}
