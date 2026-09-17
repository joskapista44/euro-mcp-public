'use strict'
const NAME=/^[A-Za-z_\\][A-Za-z0-9_.]*$/
function validName(v){return typeof v==='string'&&NAME.test(v)}
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||!['set_defined_name','rename_defined_name','delete_defined_name'].includes(op.intent))return {ok:false,outcome:'xlsx-defined-name-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(!validName(op.name))return {ok:false,outcome:'xlsx-defined-name-task-invalid-name',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='delete_defined_name')return {ok:true,outcome:'xlsx-defined-name-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,name:op.name}}
 if(typeof op.refersTo!=='string'||!op.refersTo.startsWith('='))return {ok:false,outcome:'xlsx-defined-name-task-invalid-reference',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='rename_defined_name'&&(!validName(op.newName)||op.newName===op.name))return {ok:false,outcome:'xlsx-defined-name-task-invalid-new-name',authority:'PLAN_ONLY',writeAllowed:false}
 const operation={index:0,intent:op.intent,name:op.name,refersTo:op.refersTo};if(op.intent==='rename_defined_name')operation.newName=op.newName
 return {ok:true,outcome:'xlsx-defined-name-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation}
}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeDefinedNameTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-defined-name-task-initial-read-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.definedNameObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-defined-name-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 const target={name:op.name};if(op.newName)target.newName=op.newName
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-defined-name-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:target,identityProof:'fresh-live-defined-name-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-defined-name-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.definedNameObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.state)!==JSON.stringify(pre.state))return {ok:false,outcome:'xlsx-defined-name-task-state-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.definedNameObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-defined-name-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-defined-name-task-final-read-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.definedNameObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-defined-name-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-defined-name-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:target,identityProof:'fresh-defined-name-fingerprint+exact-public-getter-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={NAME,validName,planTask,measured,executeDefinedNameTask}
