'use strict'
const {parseA1Range}=require('./range-reader.cjs')
const NAME=/^[A-Za-z_][A-Za-z0-9_. -]*$/
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||!['set_chart','rename_chart','delete_chart'].includes(op.intent))return {ok:false,outcome:'xlsx-chart-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.sheet!=='string'||!op.sheet.trim()||typeof op.name!=='string'||!NAME.test(op.name))return {ok:false,outcome:'xlsx-chart-task-invalid-identity',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='delete_chart')return {ok:true,outcome:'xlsx-chart-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,sheet:op.sheet,name:op.name}}
 if(op.intent==='rename_chart'&&(typeof op.newName!=='string'||!NAME.test(op.newName)||op.newName===op.name))return {ok:false,outcome:'xlsx-chart-task-invalid-rename-target',authority:'PLAN_ONLY',writeAllowed:false}
 const range=op.intent==='set_chart'?parseA1Range(op.range):null
 if((op.intent==='set_chart'&&!range)||typeof op.chartType!=='string'||!op.chartType||typeof op.title!=='string'||!op.title||!Number.isInteger(op.expectedSeriesCount)||op.expectedSeriesCount<1)return {ok:false,outcome:'xlsx-chart-task-invalid-set-state',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='rename_chart'&&(op.width==null||op.height==null))return {ok:false,outcome:'xlsx-chart-task-rename-fingerprint-required',authority:'PLAN_ONLY',writeAllowed:false}
 const width=op.width==null?3600000:Number(op.width),height=op.height==null?2160000:Number(op.height)
 if(!Number.isFinite(width)||width<=0||!Number.isFinite(height)||height<=0)return {ok:false,outcome:'xlsx-chart-task-invalid-size',authority:'PLAN_ONLY',writeAllowed:false}
 const presentation=op.presentation==null?null:op.presentation
 if(presentation!==null){
  if(typeof presentation!=='object'||Array.isArray(presentation))return {ok:false,outcome:'xlsx-chart-task-invalid-presentation',authority:'PLAN_ONLY',writeAllowed:false}
  const allowed=['legendPosition','horizontalAxisTitle','verticalAxisTitle','dataLabels','style']
  if(Object.keys(presentation).some(k=>!allowed.includes(k))||!Object.keys(presentation).length)return {ok:false,outcome:'xlsx-chart-task-invalid-presentation',authority:'PLAN_ONLY',writeAllowed:false}
  for(const k of ['legendPosition','horizontalAxisTitle','verticalAxisTitle'])if(presentation[k]!=null&&(typeof presentation[k]!=='string'||!presentation[k]))return {ok:false,outcome:'xlsx-chart-task-invalid-presentation',authority:'PLAN_ONLY',writeAllowed:false,field:k}
  if(presentation.style!=null&&(!Number.isInteger(presentation.style)||presentation.style<0))return {ok:false,outcome:'xlsx-chart-task-invalid-presentation',authority:'PLAN_ONLY',writeAllowed:false,field:'style'}
  if(presentation.dataLabels!=null){const d=presentation.dataLabels;if(typeof d!=='object'||Array.isArray(d)||Object.keys(d).some(k=>!['showSeriesName','showCategoryName','showValue','showPercent'].includes(k))||Object.keys(d).length!==4||Object.values(d).some(v=>typeof v!=='boolean'))return {ok:false,outcome:'xlsx-chart-task-invalid-data-labels',authority:'PLAN_ONLY',writeAllowed:false}}
 }
 const position=op.chartPosition==null?null:op.chartPosition
 if(position!==null){const keys=['fromCol','colOffset','fromRow','rowOffset'];if(typeof position!=='object'||Array.isArray(position)||Object.keys(position).some(k=>!keys.includes(k))||keys.some(k=>!Number.isFinite(position[k])))return {ok:false,outcome:'xlsx-chart-task-invalid-position',authority:'PLAN_ONLY',writeAllowed:false}}
 const series=op.series==null?null:op.series
 if(series!==null){
  if(!Array.isArray(series)||!series.length)return {ok:false,outcome:'xlsx-chart-task-invalid-series',authority:'PLAN_ONLY',writeAllowed:false}
  const seen=new Set()
  for(const s of series){const keys=['index','name','valuesRange','xValuesRange','categoryRange'];if(!s||typeof s!=='object'||Array.isArray(s)||Object.keys(s).some(k=>!keys.includes(k))||!Number.isInteger(s.index)||s.index<0||s.index>=op.expectedSeriesCount||seen.has(s.index)||Object.keys(s).length<2)return {ok:false,outcome:'xlsx-chart-task-invalid-series',authority:'PLAN_ONLY',writeAllowed:false};seen.add(s.index);for(const k of keys.slice(1))if(s[k]!=null&&(typeof s[k]!=='string'||!s[k]))return {ok:false,outcome:'xlsx-chart-task-invalid-series',authority:'PLAN_ONLY',writeAllowed:false,field:k}}
 }
 return {ok:true,outcome:'xlsx-chart-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,sheet:op.sheet,name:op.name,newName:op.intent==='rename_chart'?op.newName:undefined,range:range?.address,chartType:op.chartType,title:String(op.title),width,height,expectedSeriesCount:op.expectedSeriesCount,inRows:!!op.inRows,presentation:presentation?JSON.parse(JSON.stringify(presentation)):null,chartPosition:position?{fromCol:position.fromCol,colOffset:position.colOffset,fromRow:position.fromRow,rowOffset:position.rowOffset}:null,series:series?series.map(s=>({index:s.index,...(s.name==null?{}:{name:s.name}),...(s.valuesRange==null?{}:{valuesRange:s.valuesRange.replace(/^=/,'')}),...(s.xValuesRange==null?{}:{xValuesRange:s.xValuesRange.replace(/^=/,'')}),...(s.categoryRange==null?{}:{categoryRange:s.categoryRange.replace(/^=/,'')})})):null}}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeChartTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-chart-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.chartObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-chart-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 const resolvedName=op.intent==='rename_chart'?op.newName:op.name
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-chart-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{sheet:op.sheet,name:resolvedName},identityProof:'fresh-live-inventory+unique-chart-name+semantic-chart-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ'||!identity(freshInventory,op.sheet))return {ok:false,outcome:'xlsx-chart-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.chartObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.state)!==JSON.stringify(pre.state))return {ok:false,outcome:'xlsx-chart-task-state-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.chartObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-chart-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-chart-task-final-read-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.chartObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-chart-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-chart-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,name:resolvedName},identityProof:'fresh-live-inventory+unique-chart-name+semantic-chart-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={NAME,planTask,identity,measured,executeChartTask}
