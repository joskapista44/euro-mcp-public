'use strict'
const {parseA1Range}=require('./range-reader.cjs')
function clone(v){return JSON.parse(JSON.stringify(v))}
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||op.intent!=='sort_range')return {ok:false,outcome:'xlsx-sort-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.sheet!=='string'||!op.sheet.trim())return {ok:false,outcome:'xlsx-sort-task-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
 const range=parseA1Range(op.range),key=parseA1Range(op.keyRange)
 if(!range||!key||key.start.column!==key.end.column||key.start.column<range.start.column||key.end.column>range.end.column||key.start.row!==range.start.row||key.end.row!==range.end.row)return {ok:false,outcome:'xlsx-sort-task-invalid-key-range',authority:'PLAN_ONLY',writeAllowed:false}
 const order=op.order||'asc';if(!['asc','desc'].includes(order))return {ok:false,outcome:'xlsx-sort-task-invalid-order',authority:'PLAN_ONLY',writeAllowed:false}
 return {ok:true,outcome:'xlsx-sort-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:'sort_range',sheet:op.sheet,range:range.address,keyRange:key.address,order,hasHeaders:op.hasHeaders!==false}}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeSortTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-sort-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.sortObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:'xlsx-sort-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-sort-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{sheet:op.sheet,range:op.range,keyRange:op.keyRange},identityProof:'fresh-live-inventory+ordered-key-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',keys:pre.verification.keys}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ'||!identity(freshInventory,op.sheet))return {ok:false,outcome:'xlsx-sort-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.sortObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.matrix)!==JSON.stringify(pre.matrix))return {ok:false,outcome:'xlsx-sort-task-range-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.sortObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-sort-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ'||!identity(finalInventory,op.sheet))return {ok:false,outcome:'xlsx-sort-task-final-identity-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.sortObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-sort-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-sort-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,range:op.range,keyRange:op.keyRange},identityProof:'fresh-live-inventory+range-fingerprint+ordered-key-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',keys:clone(final.verification.keys)}}
}
module.exports={planTask,identity,measured,executeSortTask}
