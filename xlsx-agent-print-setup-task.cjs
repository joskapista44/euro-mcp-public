'use strict'
const setup=require('./xlsx-persistent-print-setup.cjs')
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||op.intent!=='set_print_setup'||!setup.validSpec(op))return {ok:false,outcome:'xlsx-print-setup-task-invalid',authority:'PLAN_ONLY',writeAllowed:false}
 return {ok:true,outcome:'xlsx-print-setup-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{...op,index:0}}
}
function measured(r){return r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executePrintSetupTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan;const op=plan.operation
 const pre=await api.printSetupObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:'xlsx-print-setup-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-print-setup-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const applied=await api.printSetupObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-print-setup-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const final=await api.printSetupObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-print-setup-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-print-setup-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={planTask,measured,executePrintSetupTask}
