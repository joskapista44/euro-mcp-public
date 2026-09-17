'use strict'
function filterObserveCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function normAddr(v){var s=String(v==null?'':v).replace(/\$/g,'').toUpperCase(),i=s.lastIndexOf('!');return i>=0?s.slice(i+1):s}
 function normCriterion(v){if(Array.isArray(v))return v.map(normCriterion);if(v==null)return null;var s=String(v);return s.length>1&&s[0]==='='?s.slice(1):s}
 function same(a,b){return JSON.stringify(normCriterion(a))===JSON.stringify(normCriterion(b))}
 function read(sh){
  var af=has(sh,'GetAutoFilter')?sh.GetAutoFilter():null,rg=af&&has(af,'GetRange')?af.GetRange():null,raw=af&&has(af,'GetFilters')?af.GetFilters():[]
  if(!Array.isArray(raw))return {measurable:false,reason:'filters-not-array'}
  var filters=[]
  for(var i=0;i<raw.length;i++){
   var f=raw[i],internal=f&&f.filter,colId=internal&&Number.isInteger(internal.ColId)?internal.ColId:null
   var on=has(f,'GetOn')?f.GetOn():f&&f.On
   if(on!==true)continue
   if(colId===null)return {measurable:false,reason:'active-filter-field-unavailable',arrayIndex:i}
   filters.push({field:colId+1,operator:has(f,'GetOperator')?f.GetOperator():f.Operator,criteria1:has(f,'GetCriteria1')?f.GetCriteria1():f.Criteria1,criteria2:has(f,'GetCriteria2')?f.GetCriteria2():f.Criteria2,on:true})
  }
  return {measurable:true,present:!!rg,range:rg&&has(rg,'GetAddress')?normAddr(rg.GetAddress()):null,filterMode:af&&has(af,'GetFilterMode')?af.GetFilterMode():false,filters:filters}
 }
 function matches(state){
  if(!state.measurable||!state.present||state.range!==normAddr(spec.range)||state.filters.length!==1)return false
  var f=state.filters[0];return f.field===spec.field&&String(f.operator)==String(spec.operator)&&same(f.criteria1,spec.criteria1)&&same(f.criteria2,spec.criteria2==null?null:spec.criteria2)
 }
 try{
  if(!spec||typeof spec.sheet!=='string'||typeof spec.range!=='string'||!Number.isInteger(spec.field)||spec.field<1||spec.criteria1==null)return {ok:false,outcome:'invalid-filter-spec',source:'live-coedit-editor'}
  var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null,r=sh&&has(sh,'GetRange')?sh.GetRange(spec.range):null
  if(!r||!has(r,'SetAutoFilter'))return {ok:false,outcome:'filter-api-unavailable',source:'live-coedit-editor'}
  var before=read(sh)
  if(!before.measurable)return {ok:false,outcome:'filter-state-unverifiable',source:'live-coedit-editor',state:before}
  if(before.present&&before.range!==normAddr(spec.range))return {ok:false,outcome:'filter-range-conflict',source:'live-coedit-editor',state:before}
  var satisfied=matches(before)
  if(!spec.apply)return {ok:true,outcome:satisfied?'filter-already-satisfied':'filter-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
  if(satisfied)return {ok:true,outcome:'filter-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
  if(before.filters.length&&!(before.filters.length===1&&before.filters[0].field===spec.field))return {ok:false,outcome:'filter-other-field-conflict',source:'live-coedit-editor',state:before}
  r.SetAutoFilter(spec.field,spec.criteria1,spec.operator,spec.criteria2==null?undefined:spec.criteria2)
  var after=read(sh)
  if(!after.measurable)return {ok:false,outcome:'filter-post-state-unverifiable',source:'live-coedit-editor',before:before,state:after}
  var pass=matches(after)
  return {ok:pass,outcome:pass?'filter-live-verified':'filter-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:after,verification:{measurable:true,match:pass}}
 }catch(e){return {ok:false,outcome:'filter-operation-error',source:'live-coedit-editor',error:String(e&&e.message||e)}}
}
module.exports={filterObserveCommand}
