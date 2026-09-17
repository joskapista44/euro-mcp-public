'use strict'
function validationObserveCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function normAddr(v){var s=String(v==null?'':v).replace(/\$/g,'').toUpperCase(),i=s.lastIndexOf('!');return i>=0?s.slice(i+1):s}
 function normFormula(v){return v==null?'':String(v)}
 var fields=['Type','AlertStyle','Operator','Formula1','Formula2','IgnoreBlank','InCellDropdown','InputMessage','InputTitle','ShowError','ShowInput','ErrorMessage','ErrorTitle']
 function read(range){
  if(!range||!has(range,'GetAddress')||!has(range,'GetValidation'))return {measurable:false,reason:'validation-read-api-unavailable'}
  var validation=range.GetValidation();if(!validation)return {measurable:false,reason:'validation-object-unavailable'}
  var state={measurable:true,address:normAddr(range.GetAddress())},missing=[]
  for(var i=0;i<fields.length;i++){
   var name=fields[i],getter='Get'+name,key=name.charAt(0).toLowerCase()+name.slice(1)
   if(!has(validation,getter)){missing.push(getter);continue}
   try{state[key]=validation[getter]()}catch(e){return {measurable:false,reason:'validation-getter-error',getter:getter,error:String(e&&e.message||e)}}
  }
  if(missing.length)return {measurable:false,reason:'validation-getters-unavailable',missing:missing}
  state.formula1=normFormula(state.formula1);state.formula2=normFormula(state.formula2)
  state.absent=state.type==null||state.type==='xlValidateInputOnly';state.present=!state.absent
  return state
 }
 function expected(){return {type:spec.validationType,alertStyle:spec.alertStyle,operator:spec.operator,formula1:normFormula(spec.formula1),formula2:normFormula(spec.formula2),ignoreBlank:spec.ignoreBlank,inCellDropdown:spec.inCellDropdown,inputMessage:spec.inputMessage,inputTitle:spec.inputTitle,showError:spec.showError,showInput:spec.showInput,errorMessage:spec.errorMessage,errorTitle:spec.errorTitle}}
 function setMismatches(state){
  var out=[]
  if(!state.measurable)return ['unmeasurable']
  if(state.address!==normAddr(spec.range))out.push('address')
  if(state.absent)out.push('absent')
  var want=expected(),keys=Object.keys(want)
  for(var i=0;i<keys.length;i++){var k=keys[i];if(state[k]!==want[k])out.push(k)}
  return out
 }
 function matchesSet(state){return setMismatches(state).length===0}
 function matchesClear(state){return !!state.measurable&&state.address===normAddr(spec.range)&&state.absent===true}
 try{
  if(!spec||typeof spec.sheet!=='string'||typeof spec.range!=='string'||(spec.intent!=='set_validation'&&spec.intent!=='clear_validation'))return {ok:false,outcome:'invalid-validation-spec',source:'live-coedit-editor'}
  var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null,range=sh&&has(sh,'GetRange')?sh.GetRange(spec.range):null
  if(!range)return {ok:false,outcome:'validation-api-unavailable',source:'live-coedit-editor'}
  var before=read(range)
  if(!before.measurable)return {ok:false,outcome:'validation-state-unverifiable',source:'live-coedit-editor',state:before}
  if(before.address!==normAddr(spec.range))return {ok:false,outcome:'validation-range-identity-mismatch',source:'live-coedit-editor',state:before}
  var satisfied=spec.intent==='set_validation'?matchesSet(before):matchesClear(before)
  var beforeMismatches=spec.intent==='set_validation'?setMismatches(before):(satisfied?[]:['validation-present'])
  if(!spec.apply)return {ok:true,outcome:satisfied?'validation-already-satisfied':'validation-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied,mismatches:beforeMismatches}}
  if(satisfied)return {ok:true,outcome:'validation-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true,mismatches:[]}}
  var validation=range.GetValidation()
  if(spec.intent==='clear_validation'){
   if(!has(validation,'Delete'))return {ok:false,outcome:'validation-delete-api-unavailable',source:'live-coedit-editor',state:before}
   validation.Delete()
  }else{
   var method=before.absent?'Add':'Modify'
   if(!has(validation,method))return {ok:false,outcome:'validation-'+method.toLowerCase()+'-api-unavailable',source:'live-coedit-editor',state:before}
   validation[method](spec.validationType,spec.alertStyle,spec.operator,spec.formula1,spec.formula2)
   var setters={ignoreBlank:'SetIgnoreBlank',inCellDropdown:'SetInCellDropdown',inputMessage:'SetInputMessage',inputTitle:'SetInputTitle',showError:'SetShowError',showInput:'SetShowInput',errorMessage:'SetErrorMessage',errorTitle:'SetErrorTitle'}
   validation=range.GetValidation()
   for(var k in setters){var setter=setters[k];if(!has(validation,setter))return {ok:false,outcome:'validation-option-api-unavailable',source:'live-coedit-editor',option:k,state:before};validation[setter](spec[k])}
  }
  var after=read(range)
  if(!after.measurable)return {ok:false,outcome:'validation-post-state-unverifiable',source:'live-coedit-editor',before:before,state:after}
  var pass=spec.intent==='set_validation'?matchesSet(after):matchesClear(after)
  var afterMismatches=spec.intent==='set_validation'?setMismatches(after):(pass?[]:['validation-present'])
  return {ok:pass,outcome:pass?'validation-live-verified':'validation-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:after,verification:{measurable:true,match:pass,mismatches:afterMismatches}}
 }catch(e){return {ok:false,outcome:'validation-operation-error',source:'live-coedit-editor',error:String(e&&e.message||e)}}
}
module.exports={validationObserveCommand}
