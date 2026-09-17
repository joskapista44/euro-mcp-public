'use strict'
const {parseA1Range}=require('./range-reader.cjs')
const NAME=/^[A-Za-z_][A-Za-z0-9_. -]*$/
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||!['set_chart','delete_chart'].includes(op.intent))return {ok:false,outcome:'xlsx-chart-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.sheet!=='string'||!op.sheet.trim()||typeof op.name!=='string'||!NAME.test(op.name))return {ok:false,outcome:'xlsx-chart-task-invalid-identity',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='delete_chart')return {ok:true,outcome:'xlsx-chart-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,sheet:op.sheet,name:op.name}}
 const range=parseA1Range(op.range)
 if(!range||typeof op.chartType!=='string'||!op.chartType||typeof op.title!=='string'||!op.title||!Number.isInteger(op.expectedSeriesCount)||op.expectedSeriesCount<1)return {ok:false,outcome:'xlsx-chart-task-invalid-set-state',authority:'PLAN_ONLY',writeAllowed:false}
 const width=op.width==null?3600000:Number(op.width),height=op.height==null?2160000:Number(op.height)
 if(!Number.isFinite(width)||width<=0||!Number.isFinite(height)||height<=0)return {ok:false,outcome:'xlsx-chart-task-invalid-size',authority:'PLAN_ONLY',writeAllowed:false}
 return {ok:true,outcome:'xlsx-chart-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,sheet:op.sheet,name:op.name,range:range.address,chartType:op.chartType,title:String(op.title),width,height,expectedSeriesCount:op.expectedSeriesCount,inRows:!!op.inRows}}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeChartTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-chart-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.chartObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-chart-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-chart-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{sheet:op.sheet,name:op.name},identityProof:'fresh-live-inventory+unique-chart-name+semantic-chart-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
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
 return {ok:true,outcome:'xlsx-chart-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,name:op.name},identityProof:'fresh-live-inventory+unique-chart-name+semantic-chart-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={NAME,planTask,identity,measured,executeChartTask}
