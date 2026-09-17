'use strict'
function pivotObserveCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function normAddr(v){var s=String(v==null?'':v).replace(/\$/g,'').toUpperCase(),i=s.lastIndexOf('!');return i>=0?s.slice(i+1):s}
 function safe(fn){try{return {ok:true,value:fn()}}catch(e){return {ok:false,error:String(e&&e.message||e)}}}
 function pivot(name){if(!has(Api,'GetPivotByName'))return null;try{return Api.GetPivotByName(name)}catch(_){return null}}
 function scalar(v){return v==null?null:String(v)}
 function snap(p){
  if(!p)return {measurable:true,present:false,name:spec.name}
  var required=['GetName','GetSource','GetParent','GetRowFields','GetColumnFields','GetDataFields','GetStyleName','GetData']
  for(var i=0;i<required.length;i++)if(!has(p,required[i]))return {measurable:false,reason:'pivot-getter-unavailable',getter:required[i]}
  var src=safe(function(){return p.GetSource()}),parent=safe(function(){return p.GetParent()})
  if(!src.ok||!parent.ok||!src.value||!parent.value||!has(src.value,'GetAddress')||!has(parent.value,'GetName'))return {measurable:false,reason:'pivot-identity-unavailable'}
  var state={measurable:true,present:true,name:p.GetName(),source:normAddr(src.value.GetAddress()),parentSheet:parent.value.GetName(),rowFields:p.GetRowFields().length,columnFields:p.GetColumnFields().length,dataFields:p.GetDataFields().length,styleName:p.GetStyleName(),assertions:[]}
  var assertions=Array.isArray(spec.assertions)?spec.assertions:[]
  for(var j=0;j<assertions.length;j++){var a=assertions[j],r=safe(function(){return p.GetData(a.items)});state.assertions.push({items:a.items,ok:r.ok,value:r.ok?scalar(r.value):null,error:r.ok?null:r.error,expected:scalar(a.expected)})}
  return state
 }
 function matchCreate(state){if(!state.measurable||!state.present||state.name!==spec.name||state.source!==normAddr(spec.sourceRange)||state.styleName!==spec.styleName||state.rowFields!==1||state.columnFields!==1||state.dataFields!==1)return false;return state.assertions.length===spec.assertions.length&&state.assertions.every(function(x){return x.ok&&x.value===x.expected})}
 function deleteState(){var p=pivot(spec.name),ps=snap(p),sheet=null;try{sheet=has(Api,'GetSheet')?Api.GetSheet(spec.pivotSheet):null}catch(_){}return {pivot:ps,pivotSheetPresent:!!sheet}}
 function matchDelete(state){return state.pivot&&state.pivot.measurable&&!state.pivot.present&&state.pivotSheetPresent===false}
 try{
  if(!spec||!['create_pivot','delete_pivot_sheet'].includes(spec.intent)||typeof spec.name!=='string')return {ok:false,outcome:'invalid-pivot-spec',source:'live-coedit-editor'}
  var before=spec.intent==='create_pivot'?snap(pivot(spec.name)):deleteState(),satisfied=spec.intent==='create_pivot'?matchCreate(before):matchDelete(before)
  var measurable=spec.intent==='create_pivot'?before.measurable:before.pivot.measurable
  if(!measurable)return {ok:false,outcome:'pivot-state-unverifiable',source:'live-coedit-editor',state:before}
  if(!spec.apply)return {ok:true,outcome:satisfied?'pivot-already-satisfied':'pivot-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
  if(satisfied)return {ok:true,outcome:'pivot-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
  if(spec.intent==='create_pivot'){
   if(before.present)return {ok:false,outcome:'pivot-name-conflict',source:'live-coedit-editor',state:before}
   if(!has(Api,'GetSheet')||!has(Api,'InsertPivotNewWorksheet'))return {ok:false,outcome:'pivot-create-api-unavailable',source:'live-coedit-editor'}
   var sourceSheet=Api.GetSheet(spec.sourceSheet),source=sourceSheet&&has(sourceSheet,'GetRange')?sourceSheet.GetRange(spec.sourceRange):null
   if(!source)return {ok:false,outcome:'pivot-source-unavailable',source:'live-coedit-editor'}
   var created=Api.InsertPivotNewWorksheet(source)
   if(!created||!has(created,'SetName')||!has(created,'AddFields')||!has(created,'AddDataField')||!has(created,'SetStyleName'))return {ok:false,outcome:'pivot-build-api-unavailable',source:'live-coedit-editor'}
   created.SetName(spec.name);created.AddFields({rows:spec.rowField,columns:spec.columnField});created.AddDataField(spec.dataField);created.SetStyleName(spec.styleName);if(has(created,'RefreshTable'))created.RefreshTable()
  }else{
   if(!before.pivot.present||before.pivot.parentSheet!==spec.pivotSheet)return {ok:false,outcome:'pivot-delete-identity-conflict',source:'live-coedit-editor',state:before}
   var target=Api.GetSheet(spec.pivotSheet);if(!target||!has(target,'Delete'))return {ok:false,outcome:'pivot-sheet-delete-api-unavailable',source:'live-coedit-editor',state:before};target.Delete()
  }
  var after=spec.intent==='create_pivot'?snap(pivot(spec.name)):deleteState(),pass=spec.intent==='create_pivot'?matchCreate(after):matchDelete(after)
  return {ok:pass,outcome:pass?'pivot-live-verified':'pivot-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:after,verification:{measurable:true,match:pass}}
 }catch(e){return {ok:false,outcome:'pivot-operation-error',source:'live-coedit-editor',error:String(e&&e.message||e)}}
}
module.exports={pivotObserveCommand}
