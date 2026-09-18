'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-chart-task.cjs')
const {runChartInFrame}=require('./live-charts.cjs')
const {runChartPresentationInFrame}=require('./live-chart-presentation.cjs')
const {runChartDataObjectInFrame}=require('./live-chart-data-objects.cjs')
const {runCommand:runDefinedNameCommand}=require('./xlsx-persistent-defined-name.cjs')
function chartSemanticStateCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function safe(o,n){try{return has(o,n)?o[n]():null}catch(_){return null}}
 function trim(v){return typeof v==='string'?v.replace(/[\r\n]+$/g,''):v}
 function formula(v){return typeof v==='string'&&v.charAt(0)==='='?v.slice(1):v}
 function seriesName(v){if(typeof v!=='string')return v;if(v.slice(0,2)==='="'&&v.slice(-1)==='"')return v.slice(2,-1).replace(/""/g,'"');return v}
 try{
  var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh||!has(sh,'GetAllCharts'))return {ok:false,outcome:'chart-state-unverifiable',source:'live-coedit-editor'}
  if(has(sh,'SetActive'))sh.SetActive()
  if(has(Api,'GetActiveSheet')){var active=Api.GetActiveSheet(),activeName=safe(active,'GetName');if(activeName!=null&&String(activeName)!==String(spec.sheet))return {ok:false,outcome:'chart-sheet-activation-failed',source:'live-coedit-editor',expectedSheet:spec.sheet,actualSheet:activeName}}
  var charts=sh.GetAllCharts()||[],matches=[];for(var i=0;i<charts.length;i++)if(String(safe(charts[i],'GetName'))===String(spec.name))matches.push(charts[i])
  if(matches.length!==1)return {ok:true,outcome:'chart-state-read',source:'live-coedit-editor',state:{measurable:true,present:matches.length>0,count:matches.length,name:spec.name}}
  var c=matches[0],all=has(c,'GetAllSeries')?(c.GetAllSeries()||[]):null,unknown=[]
  var state={measurable:true,present:true,count:1,name:safe(c,'GetName'),chartType:safe(c,'GetChartType'),title:trim(safe(c,'GetTitle')),width:safe(c,'GetWidth'),height:safe(c,'GetHeight'),seriesCount:all?all.length:null}
  if(spec.geometryIdentityName){
   var marker=null;try{marker=has(Api,'GetDefName')?Api.GetDefName(spec.geometryIdentityName):null}catch(_){}
   state.geometryIdentityPresent=!!marker
   state.geometryIdentityRef=marker?safe(marker,'GetRefersTo'):null
   state.geometryIdentityVerified=!!marker&&String(safe(marker,'GetName'))===String(spec.geometryIdentityName)&&String(state.geometryIdentityRef)===String(spec.geometryIdentityRef)
   // Reopened charts can expose 0x0 through public geometry getters even
   // though SetSize was previously live-verified. The exact durable marker is
   // accepted only for that known unavailable state, never over non-zero data.
   if(state.geometryIdentityVerified&&Number(state.width)===0&&Number(state.height)===0){state.width=spec.width;state.height=spec.height;state.geometryReadback='durable-identity-fallback'}
  }
  var coreRequired=spec.intent==='set_chart'?['name','chartType','title','width','height','seriesCount']:['name'];for(var j=0;j<coreRequired.length;j++){var core=coreRequired[j];if(state[core]==null)unknown.push(core)}
  var p=spec.presentation
  if(p){
   if(p.legendPosition!=null){state.legendPosition=safe(c,'GetLegendPos');if(state.legendPosition==null)unknown.push('legendPosition')}
   if(p.horizontalAxisTitle!=null){state.horizontalAxisTitle=trim(safe(c,'GetHorAxisTitle'));if(state.horizontalAxisTitle==null)unknown.push('horizontalAxisTitle')}
   if(p.verticalAxisTitle!=null){state.verticalAxisTitle=trim(safe(c,'GetVerAxisTitle'));if(state.verticalAxisTitle==null)unknown.push('verticalAxisTitle')}
   if(p.dataLabels!=null){state.dataLabels=safe(c,'GetDataLabels');if(state.dataLabels==null)unknown.push('dataLabels')}
   if(p.style!=null){state.style=safe(c,'GetChartStyle');if(state.style==null)unknown.push('style')}
  }
  if(spec.chartPosition){state.position=safe(c,'GetPosition');if(state.position==null)unknown.push('position')}
  if(spec.series){state.series=[];for(var k=0;k<spec.series.length;k++){var wanted=spec.series[k],s=all&&all[wanted.index],actual={index:wanted.index};if(!s){unknown.push('series['+wanted.index+']');state.series.push(actual);continue}if(wanted.name!=null){actual.name=seriesName(safe(s,'GetName'));if(actual.name==null)unknown.push('seriesName['+wanted.index+']')}if(wanted.valuesRange!=null){actual.valuesRange=formula(safe(s,'GetValues'));if(actual.valuesRange==null)unknown.push('seriesValues['+wanted.index+']')}if(wanted.xValuesRange!=null){actual.xValuesRange=formula(safe(s,'GetXValues'));if(actual.xValuesRange==null)unknown.push('seriesXValues['+wanted.index+']')}if(wanted.categoryRange!=null){actual.categoryRange=formula(safe(s,'GetCatFormula'));if(actual.categoryRange==null)unknown.push('categoryFormula['+wanted.index+']')}state.series.push(actual)}}
  if(unknown.length){state.measurable=false;state.unknown=unknown;return {ok:false,outcome:'chart-state-unverifiable',source:'live-coedit-editor',state:state}}
  return {ok:true,outcome:'chart-state-read',source:'live-coedit-editor',state:state}
 }catch(e){return {ok:false,outcome:'chart-state-error',source:'live-coedit-editor',error:String(e&&e.message||e)} }
}
async function semanticState(session,op){
 const body=`return (${chartSemanticStateCommand.toString()})(${JSON.stringify(op)});`
 return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),8000);try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v)})}catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message||err)})}}),{where:session.apiWhere,body})
}
function stateOf(observed,op){
 const matches=(observed?.charts||[]).filter(c=>String(c?.name)===op.name)
 if(matches.length!==1)return {measurable:true,present:matches.length>0,count:matches.length,name:op.name}
 const c=matches[0]
 return {measurable:true,present:true,count:1,name:c.name,chartType:c.chartType,title:c.title==null?null:String(c.title).replace(/[\r\n]+$/g,''),width:c.width,height:c.height,seriesCount:c.seriesCount}
}
function chartStateMatches(state,op){
 if(!(state.measurable&&state.present&&state.count===1&&String(state.chartType)===op.chartType&&state.title===op.title&&Number(state.width)===op.width&&Number(state.height)===op.height&&Number(state.seriesCount)===op.expectedSeriesCount))return false
 if(op.geometryIdentityName&&state.geometryIdentityVerified!==true)return false
 const p=op.presentation||{}
 if(p.legendPosition!=null&&String(state.legendPosition)!==p.legendPosition)return false
 if(p.horizontalAxisTitle!=null&&state.horizontalAxisTitle!==p.horizontalAxisTitle)return false
 if(p.verticalAxisTitle!=null&&state.verticalAxisTitle!==p.verticalAxisTitle)return false
 if(p.dataLabels!=null){const expected={showSerName:p.dataLabels.showSeriesName,showCatName:p.dataLabels.showCategoryName,showVal:p.dataLabels.showValue,showPercent:p.dataLabels.showPercent};if(!state.dataLabels||Object.keys(expected).some(k=>state.dataLabels[k]!==expected[k]))return false}
 if(p.style!=null&&Number(state.style)!==p.style)return false
 if(op.chartPosition&&(!state.position||['fromCol','colOffset','fromRow','rowOffset'].some(k=>Number(state.position[k])!==Number(op.chartPosition[k]))))return false
 if(op.series&&op.series.some(w=>{const a=(state.series||[]).find(x=>x.index===w.index);return !a||Object.keys(w).some(k=>a[k]!==w[k])}))return false
 return true
}
function match(state,op){
 if(op.intent==='delete_chart')return state.measurable&&state.present===false&&state.count===0
 if(op.intent==='rename_chart')return state.source?.measurable&&state.source.present===false&&state.source.count===0&&chartStateMatches(state.target,op)
 return chartStateMatches(state,op)
}
async function operationState(session,op){
 if(op.intent!=='rename_chart')return semanticState(session,op)
 const source=await semanticState(session,{...op,intent:'set_chart',name:op.name}),target=await semanticState(session,{...op,intent:'set_chart',name:op.newName})
 if(!source?.ok||!target?.ok)return {ok:false,outcome:'chart-rename-state-unverifiable',source:'live-coedit-editor',sourceRead:source,targetRead:target}
 return {ok:true,outcome:'chart-rename-state-read',source:'live-coedit-editor',state:{measurable:true,source:source.state,target:target.state}}
}
function advancedMutationSpecs(before,op){
 const p=op.presentation||{},calls=[]
 // ApplyChartStyle resets presentation state. Build underlying style/data first,
 // then apply the requested legend/axes/labels as the final semantic layer.
 if(p.style!=null&&Number(before.style)!==p.style)calls.push({runner:'presentation',spec:{type:'chart.presentation.style',sheet:op.sheet,name:op.name,style:p.style}})
 for(const wanted of op.series||[]){const old=(before.series||[]).find(s=>s.index===wanted.index)||{};if(wanted.name!=null&&old.name!==wanted.name)calls.push({runner:'data',spec:{type:'chart.data.seriesName',sheet:op.sheet,name:op.name,seriesIndex:wanted.index,value:wanted.name}});if(wanted.valuesRange!=null&&old.valuesRange!==wanted.valuesRange)calls.push({runner:'data',spec:{type:'chart.data.seriesValues',sheet:op.sheet,name:op.name,seriesIndex:wanted.index,range:wanted.valuesRange}});if(wanted.xValuesRange!=null&&old.xValuesRange!==wanted.xValuesRange)calls.push({runner:'data',spec:{type:'chart.data.seriesXValues',sheet:op.sheet,name:op.name,seriesIndex:wanted.index,range:wanted.xValuesRange}});if(wanted.categoryRange!=null&&old.categoryRange!==wanted.categoryRange)calls.push({runner:'data',spec:{type:'chart.data.categoryFormula',sheet:op.sheet,name:op.name,seriesIndex:wanted.index,range:wanted.categoryRange}})}
 if(op.chartPosition&&(!before.position||['fromCol','colOffset','fromRow','rowOffset'].some(k=>Number(before.position[k])!==Number(op.chartPosition[k]))))calls.push({runner:'data',spec:{type:'chart.object.position',sheet:op.sheet,name:op.name,...op.chartPosition}})
 if(p.legendPosition!=null&&before.legendPosition!==p.legendPosition)calls.push({runner:'presentation',spec:{type:'chart.presentation.legend',sheet:op.sheet,name:op.name,position:p.legendPosition}})
 if((p.horizontalAxisTitle!=null&&before.horizontalAxisTitle!==p.horizontalAxisTitle)||(p.verticalAxisTitle!=null&&before.verticalAxisTitle!==p.verticalAxisTitle))calls.push({runner:'presentation',spec:{type:'chart.presentation.axisTitles',sheet:op.sheet,name:op.name,horizontal:p.horizontalAxisTitle,vertical:p.verticalAxisTitle}})
 if(p.dataLabels!=null){const expected={showSerName:p.dataLabels.showSeriesName,showCatName:p.dataLabels.showCategoryName,showVal:p.dataLabels.showValue,showPercent:p.dataLabels.showPercent};if(!before.dataLabels||Object.keys(expected).some(k=>before.dataLabels[k]!==expected[k]))calls.push({runner:'presentation',spec:{type:'chart.presentation.dataLabels',sheet:op.sheet,name:op.name,...p.dataLabels}})}
 return calls
}
async function chartObserved(session,op,apply){
 const inspect=()=>operationState(session,op)
 const beforeRead=await inspect()
 if(!beforeRead?.ok)return {ok:false,outcome:'chart-state-unverifiable',source:'live-coedit-editor',readback:beforeRead}
 const before=beforeRead.state,satisfied=match(before,op)
 if(!apply)return {ok:true,outcome:satisfied?'chart-already-satisfied':'chart-observed',source:'live-coedit-editor',noOp:satisfied,state:before,verification:{measurable:true,match:satisfied}}
 if(satisfied)return {ok:true,outcome:'chart-already-satisfied',source:'live-coedit-editor',noOp:true,state:before,verification:{measurable:true,match:true}}
 if(op.geometryIdentityName&&before.geometryIdentityPresent&&!before.geometryIdentityVerified)return {ok:false,outcome:'chart-geometry-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
 let command,mutated=false,mutations=[]
 if(op.intent==='delete_chart'){
  if(before.count!==1)return {ok:false,outcome:'chart-delete-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
  command={type:'chart.delete',sheet:op.sheet,name:op.name}
 }else if(op.intent==='rename_chart'){
  if(before.source?.count!==1||before.target?.count!==0||!chartStateMatches(before.source,op))return {ok:false,outcome:'chart-rename-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
  command={type:'chart.object.rename',sheet:op.sheet,name:op.name,newName:op.newName}
 }else if(before.count===0){
  command={type:'chart.create',sheet:op.sheet,range:op.range,chartType:op.chartType,name:op.name,title:op.title,width:op.width,height:op.height,inRows:op.inRows}
 }else if(before.count===1&&String(before.chartType)===op.chartType&&Number(before.seriesCount)===op.expectedSeriesCount){
  command=before.title!==op.title||Number(before.width)!==op.width||Number(before.height)!==op.height?{type:'chart.modify',sheet:op.sheet,name:op.name,title:op.title,width:op.width,height:op.height}:null
 }else return {ok:false,outcome:'chart-set-identity-conflict',source:'live-coedit-editor',state:before,verification:{measurable:true,match:false}}
 let changed=null
 if(command){
  const runner=op.intent==='rename_chart'?runChartDataObjectInFrame:runChartInFrame;changed=await runner(session.frame,session.apiWhere,command,8000);mutations.push(changed)
  if(op.intent==='set_chart'&&command.type==='chart.create'&&changed?.ok&&changed?.outcome==='size-pending'&&changed?.verification?.status==='PENDING'){
   const sizeRepair=await runChartInFrame(session.frame,session.apiWhere,{type:'chart.modify',sheet:op.sheet,name:op.name,title:op.title,width:op.width,height:op.height},8000);mutations.push(sizeRepair)
   if(sizeRepair?.ok&&sizeRepair?.verification?.status==='PASS')mutated=true
   else return {ok:false,outcome:'chart-mutation-unverified',source:'live-coedit-editor',state:before,mutation:sizeRepair,mutations,mutationAttempted:true,verification:{measurable:true,match:false}}
  }else{
   if(changed?.ok&&changed?.verification?.status==='PASS')mutated=true
   if(!changed?.ok||changed?.verification?.status!=='PASS')return {ok:false,outcome:'chart-mutation-unverified',source:'live-coedit-editor',state:before,mutation:changed,mutationAttempted:mutated,verification:{measurable:true,match:false}}
  }
 }
 if(op.intent==='set_chart'){
  for(const call of advancedMutationSpecs(before,op)){const fn=call.runner==='presentation'?runChartPresentationInFrame:runChartDataObjectInFrame,r=await fn(session.frame,session.apiWhere,call.spec,8000);mutations.push(r);if(r?.ok&&r?.verification?.status==='PASS')mutated=true;else return {ok:false,outcome:'chart-mutation-unverified',source:'live-coedit-editor',state:before,mutation:r,mutations,mutationAttempted:mutated,verification:{measurable:true,match:false}}}
  if(op.geometryIdentityName&&!before.geometryIdentityVerified){
   const measured=await inspect(),withoutMarker={...op,geometryIdentityName:null,geometryIdentityRef:null}
   if(!measured?.ok||!chartStateMatches(measured.state,withoutMarker))return {ok:false,outcome:'chart-geometry-not-proven-before-identity',source:'live-coedit-editor',state:measured?.state||null,mutations,mutationAttempted:mutated,verification:{measurable:!!measured?.ok,match:false}}
   const marker=await runDefinedNameCommand(session,{intent:'set_defined_name',name:op.geometryIdentityName,refersTo:op.geometryIdentityRef},true)
   mutations.push({operation:'chart.geometry.identity',...marker})
   if(!marker?.ok||marker?.noOp===true||marker?.verification?.match!==true)return {ok:false,outcome:'chart-geometry-identity-unverified',source:'live-coedit-editor',state:measured.state,mutation:marker,mutations,mutationAttempted:mutated,verification:{measurable:true,match:false}}
   mutated=true
  }
 }
 const afterRead=await inspect()
 if(!afterRead?.ok)return {ok:false,outcome:'chart-post-state-unverifiable',source:'live-coedit-editor',mutation:changed,mutations,mutationAttempted:mutated,readback:afterRead}
 const after=afterRead.state,pass=match(after,op)
 return {ok:pass,outcome:pass?'chart-live-verified':'chart-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:mutated,before,state:after,mutation:changed,mutations,mutationAttempted:mutated,verification:{measurable:true,match:pass}}
}
async function executeChartTaskInPersistentSession(options={}){
 if(!options.url||!options.user||!options.pass)return {ok:false,outcome:'xlsx-persistent-credentials-required',authority:'PLAN_ONLY',writeAllowed:false}
 return persistent.withPersistentXlsxSession(options,api=>agent.executeChartTask({task:options.task,api:{...api,chartObserved:async(op,apply)=>{const r=await chartObserved(api.session,op,apply);if(apply&&(r?.applied||r?.mutationAttempted))api.session.markWrite();return r}}}))
}
const runCommand=chartObserved
module.exports={chartSemanticStateCommand,semanticState,stateOf,chartStateMatches,match,operationState,advancedMutationSpecs,chartObserved,runCommand,executeChartTaskInPersistentSession}
