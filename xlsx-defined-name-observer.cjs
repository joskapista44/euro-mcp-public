'use strict'
function definedNameObserveCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function read(name){
  if(!has(Api,'GetDefName'))return {measurable:false,reason:'defined-name-read-api-unavailable'}
  var obj=null;try{obj=Api.GetDefName(name)}catch(_){obj=null}
  if(!obj)return {measurable:true,present:false,name:name,refersTo:null}
  if(!has(obj,'GetName')||!has(obj,'GetRefersTo'))return {measurable:false,reason:'defined-name-getters-unavailable',name:name}
  try{return {measurable:true,present:true,name:obj.GetName(),refersTo:obj.GetRefersTo()}}catch(e){return {measurable:false,reason:'defined-name-getter-error',name:name,error:String(e&&e.message||e)}}
 }
 function pair(){return {source:read(spec.name),target:read(spec.newName)}}
 function match(state){
  if(spec.intent==='set_defined_name')return state.measurable&&state.present&&state.name===spec.name&&state.refersTo===spec.refersTo
  if(spec.intent==='delete_defined_name')return state.measurable&&!state.present
  return state.source.measurable&&state.target.measurable&&!state.source.present&&state.target.present&&state.target.name===spec.newName&&state.target.refersTo===spec.refersTo
 }
 try{
  if(!spec||!['set_defined_name','rename_defined_name','delete_defined_name'].includes(spec.intent)||typeof spec.name!=='string')return {ok:false,outcome:'invalid-defined-name-spec',source:'live-coedit-editor'}
  var before=spec.intent==='rename_defined_name'?pair():read(spec.name)
  var measurable=spec.intent==='rename_defined_name'?before.source.measurable&&before.target.measurable:before.measurable
  if(!measurable)return {ok:false,outcome:'defined-name-state-unverifiable',source:'live-coedit-editor',state:before}
  var satisfied=match(before)
  if(!spec.apply)return {ok:true,outcome:satisfied?'defined-name-already-satisfied':'defined-name-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
  if(satisfied)return {ok:true,outcome:'defined-name-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
  if(spec.intent==='set_defined_name'){
   if(before.present){var existing=Api.GetDefName(spec.name);if(!has(existing,'SetRefersTo'))return {ok:false,outcome:'defined-name-set-api-unavailable',source:'live-coedit-editor',state:before};existing.SetRefersTo(spec.refersTo)}
   else{if(!has(Api,'AddDefName'))return {ok:false,outcome:'defined-name-add-api-unavailable',source:'live-coedit-editor',state:before};Api.AddDefName(spec.name,spec.refersTo)}
  }else if(spec.intent==='delete_defined_name'){
   var doomed=Api.GetDefName(spec.name);if(!has(doomed,'Delete'))return {ok:false,outcome:'defined-name-delete-api-unavailable',source:'live-coedit-editor',state:before};doomed.Delete()
  }else{
   if(!before.source.present||before.source.refersTo!==spec.refersTo)return {ok:false,outcome:'defined-name-source-conflict',source:'live-coedit-editor',state:before}
   if(before.target.present)return {ok:false,outcome:'defined-name-target-conflict',source:'live-coedit-editor',state:before}
   var source=Api.GetDefName(spec.name);if(!has(source,'SetName'))return {ok:false,outcome:'defined-name-rename-api-unavailable',source:'live-coedit-editor',state:before};source.SetName(spec.newName)
  }
  var after=spec.intent==='rename_defined_name'?pair():read(spec.name),pass=match(after)
  return {ok:pass,outcome:pass?'defined-name-live-verified':'defined-name-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:after,verification:{measurable:true,match:pass}}
 }catch(e){return {ok:false,outcome:'defined-name-operation-error',source:'live-coedit-editor',error:String(e&&e.message||e)}}
}
module.exports={definedNameObserveCommand}
