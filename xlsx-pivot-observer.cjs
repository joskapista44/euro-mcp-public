'use strict'
function pivotObserveCommand(spec){
 var stage='init'
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function normAddr(v){var s=String(v==null?'':v).replace(/\$/g,'').toUpperCase(),i=s.lastIndexOf('!');return i>=0?s.slice(i+1):s}
 function absoluteRef(sheet,range){var parts=normAddr(range).split(':');return '='+sheet+'!'+parts.map(function(cell){var m=cell.match(/^([A-Z]+)([0-9]+)$/);return m?'$'+m[1]+'$'+m[2]:cell}).join(':')}
 function safe(fn){try{return {ok:true,value:fn()}}catch(e){return {ok:false,error:String(e&&e.message||e)}}}
 function pivot(name){if(!has(Api,'GetPivotByName'))return null;try{return Api.GetPivotByName(name)}catch(_){return null}}
 function scalar(v){return v==null?null:String(v)}
 function snap(p){
  if(!p)return {measurable:true,present:false,name:spec.name}
  var required=['GetName','GetSource','GetParent','GetRowFields','GetColumnFields','GetDataFields','GetStyleName','GetData']
  for(var i=0;i<required.length;i++)if(!has(p,required[i]))return {measurable:false,reason:'pivot-getter-unavailable',getter:required[i]}
  var src=safe(function(){return p.GetSource()}),parent=safe(function(){return p.GetParent()})
  if(!parent.ok||!parent.value||!has(parent.value,'GetName'))return {measurable:false,reason:'pivot-parent-identity-unavailable'}
  var source=null,sourceSheet=null,sourceProof=null
  if(src.ok&&src.value&&has(src.value,'GetAddress')){
   source=normAddr(src.value.GetAddress());sourceProof='public-GetSource'
   if(has(src.value,'GetWorksheet')){var sourceParent=safe(function(){return src.value.GetWorksheet()});if(sourceParent.ok&&sourceParent.value&&has(sourceParent.value,'GetName'))sourceSheet=sourceParent.value.GetName()}
  }else if(spec.sourceIdentityName&&has(Api,'GetDefName')){
   var marker=safe(function(){return Api.GetDefName(spec.sourceIdentityName)}),expectedRef=absoluteRef(spec.sourceSheet,spec.sourceRange)
   if(!marker.ok||!marker.value||!has(marker.value,'GetName')||!has(marker.value,'GetRefersTo'))return {measurable:false,reason:'pivot-source-marker-unavailable',sourceError:src.error||null}
   var markerName=safe(function(){return marker.value.GetName()}),markerRef=safe(function(){return marker.value.GetRefersTo()})
   if(!markerName.ok||!markerRef.ok||markerName.value!==spec.sourceIdentityName||markerRef.value!==expectedRef)return {measurable:false,reason:'pivot-source-marker-mismatch',sourceError:src.error||null,expectedRef:expectedRef,actualRef:markerRef.ok?markerRef.value:null}
   source=normAddr(spec.sourceRange);sourceSheet=spec.sourceSheet;sourceProof='defined-name-fallback'
  }else return {measurable:false,reason:'pivot-source-identity-unavailable',sourceError:src.error||null}
  var state={measurable:true,present:true,name:p.GetName(),source:source,sourceSheet:sourceSheet,sourceProof:sourceProof,parentSheet:parent.value.GetName(),rowFields:p.GetRowFields().length,columnFields:p.GetColumnFields().length,dataFields:p.GetDataFields().length,styleName:p.GetStyleName(),assertions:[]}
  var assertions=Array.isArray(spec.assertions)?spec.assertions:[]
  for(var j=0;j<assertions.length;j++){var a=assertions[j],r=safe(function(){return p.GetData(a.items)});state.assertions.push({items:a.items,ok:r.ok,value:r.ok?scalar(r.value):null,error:r.ok?null:r.error,expected:scalar(a.expected)})}
  return state
 }
 function matchIdentity(state){return !!state&&state.measurable&&state.present&&state.name===spec.name&&state.source===normAddr(spec.sourceRange)&&(!spec.pivotSheet||state.parentSheet===spec.pivotSheet)&&state.styleName===spec.styleName&&state.rowFields===1&&state.columnFields===1&&state.dataFields===1}
 function sameDedicatedIdentity(state){return !!state&&state.measurable&&state.present&&state.name===spec.name&&state.source===normAddr(spec.sourceRange)&&state.sourceSheet===spec.sourceSheet&&state.parentSheet===spec.pivotSheet}
 function matchCreate(state){if(!matchIdentity(state))return false;return state.assertions.length===spec.assertions.length&&state.assertions.every(function(x){return x.ok&&x.value===x.expected})}
 function deleteState(){var p=pivot(spec.name),ps=snap(p),sheet=null;try{sheet=has(Api,'GetSheet')?Api.GetSheet(spec.pivotSheet):null}catch(_){}return {pivot:ps,pivotSheetPresent:!!sheet}}
 function matchDelete(state){return state.pivot&&state.pivot.measurable&&!state.pivot.present&&state.pivotSheetPresent===false}
 try{
  if(!spec||!['create_pivot','refresh_pivot','delete_pivot_sheet'].includes(spec.intent)||typeof spec.name!=='string')return {ok:false,outcome:'invalid-pivot-spec',source:'live-coedit-editor'}
  var isPivotState=spec.intent!=='delete_pivot_sheet'
  var before=isPivotState?snap(pivot(spec.name)):deleteState(),satisfied=isPivotState?matchCreate(before):matchDelete(before)
  var measurable=isPivotState?before.measurable:before.pivot.measurable
  if(!measurable)return {ok:false,outcome:'pivot-state-unverifiable',source:'live-coedit-editor',state:before}
  if(!spec.apply)return {ok:true,outcome:satisfied?'pivot-already-satisfied':'pivot-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
  if(satisfied)return {ok:true,outcome:'pivot-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
  if(spec.intent==='create_pivot'){
   if(before.present){
    if(spec.repairIncompletePivot!==true||!spec.pivotSheet||!sameDedicatedIdentity(before))return {ok:false,outcome:'pivot-name-conflict',source:'live-coedit-editor',state:before}
    if(!has(Api,'GetAllPivotTables')||!has(Api,'GetSheet')||!has(Api,'AddSheet'))return {ok:false,outcome:'pivot-repair-api-unavailable',source:'live-coedit-editor',state:before}
    var all=Api.GetAllPivotTables()||[],onTarget=[]
    for(var q=0;q<all.length;q++){var qp=safe(function(){return all[q].GetParent().GetName()});if(qp.ok&&qp.value===spec.pivotSheet)onTarget.push(all[q])}
    if(onTarget.length!==1||!has(onTarget[0],'GetName')||onTarget[0].GetName()!==spec.name)return {ok:false,outcome:'pivot-repair-sheet-not-dedicated',source:'live-coedit-editor',state:before,pivotCount:onTarget.length}
    var staleSheet=Api.GetSheet(spec.pivotSheet)
    if(!staleSheet||!has(staleSheet,'Delete'))return {ok:false,outcome:'pivot-repair-sheet-delete-unavailable',source:'live-coedit-editor',state:before}
    stage='repair-delete-dedicated-sheet';staleSheet.Delete();stage='repair-recreate-dedicated-sheet';Api.AddSheet(spec.pivotSheet)
    if(!Api.GetSheet(spec.pivotSheet))return {ok:false,outcome:'pivot-repair-sheet-recreate-failed',source:'live-coedit-editor',state:before}
   }
   if(!has(Api,'GetRange')||(spec.pivotSheet?!has(Api,'InsertPivotExistingWorksheet'):!has(Api,'InsertPivotNewWorksheet')))return {ok:false,outcome:'pivot-create-api-unavailable',source:'live-coedit-editor'}
   // ONLYOFFICE documents pivot creation with a workbook-qualified Api.GetRange.
   // Preserve worksheet identity in the reference instead of relying on the
   // current active sheet while InsertPivotNewWorksheet builds its cache.
   var qualified="'"+spec.sourceSheet.replace(/'/g,"''")+"'!$"+spec.sourceRange.replace(/\$/g,'').replace(':',':$').replace(/([A-Z]+)([0-9]+)/g,'$1$$$2')
   stage='resolve-qualified-source';var source=Api.GetRange(qualified)
   if(!source)return {ok:false,outcome:'pivot-source-unavailable',source:'live-coedit-editor',qualifiedSource:qualified}
   var created=null
   if(spec.pivotSheet){
    if(!has(Api,'GetSheet')||!has(Api,'InsertPivotExistingWorksheet'))return {ok:false,outcome:'pivot-existing-create-api-unavailable',source:'live-coedit-editor'}
    var destinationSheet=Api.GetSheet(spec.pivotSheet),destination=destinationSheet&&has(destinationSheet,'GetRange')?destinationSheet.GetRange(spec.destinationRange):null
    if(!destination)return {ok:false,outcome:'pivot-destination-unavailable',source:'live-coedit-editor'}
    stage='insert-pivot-existing';created=Api.InsertPivotExistingWorksheet(source,destination)
   }else{stage='insert-pivot-new';created=Api.InsertPivotNewWorksheet(source)}
   if(!created||!has(created,'SetName')||!has(created,'AddFields')||!has(created,'AddDataField')||!has(created,'SetStyleName'))return {ok:false,outcome:'pivot-build-api-unavailable',source:'live-coedit-editor'}
   // AddDataField first performs asc_addDataField, then its public return-value
   // wrapper dereferences asc_getDataFields().length. Existing-sheet co-editing
   // can return null at that second step even though the mutation ran. Catch
   // only that measured post-mutation wrapper defect; the mandatory refresh and
   // complete semantic readback below still fail closed unless the field exists.
   stage='set-name';created.SetName(spec.name);stage='add-fields';created.AddFields({rows:spec.rowField,columns:spec.columnField});stage='add-data-field';try{created.AddDataField(spec.dataField)}catch(addError){var addMessage=String(addError&&addError.message||addError);if(!/null.*length/i.test(addMessage))throw addError}stage='set-style';created.SetStyleName(spec.styleName);if(has(created,'RefreshTable')){stage='refresh';created.RefreshTable()}
  }else if(spec.intent==='refresh_pivot'){
   var existing=pivot(spec.name)
   if(!matchIdentity(before))return {ok:false,outcome:'pivot-refresh-identity-conflict',source:'live-coedit-editor',state:before}
   if(!existing||!has(existing,'RefreshTable'))return {ok:false,outcome:'pivot-refresh-api-unavailable',source:'live-coedit-editor',state:before}
   existing.RefreshTable()
  }else{
   if(!before.pivot.present||before.pivot.parentSheet!==spec.pivotSheet)return {ok:false,outcome:'pivot-delete-identity-conflict',source:'live-coedit-editor',state:before}
   var target=Api.GetSheet(spec.pivotSheet);if(!target||!has(target,'Delete'))return {ok:false,outcome:'pivot-sheet-delete-api-unavailable',source:'live-coedit-editor',state:before};target.Delete()
  }
  var after=isPivotState?snap(pivot(spec.name)):deleteState(),pass=isPivotState?matchCreate(after):matchDelete(after)
  return {ok:pass,outcome:pass?'pivot-live-verified':'pivot-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:after,verification:{measurable:true,match:pass}}
 }catch(e){return {ok:false,outcome:'pivot-operation-error',source:'live-coedit-editor',stage:stage,error:String(e&&e.message||e)}}
}
module.exports={pivotObserveCommand}
