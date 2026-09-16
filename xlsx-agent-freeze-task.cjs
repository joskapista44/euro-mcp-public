'use strict'
const {parseA1Range}=require('./range-reader.cjs')
function planTask(task){
  const op=task?.operations?.length===1?task.operations[0]:null
  if(!op||!['freeze_panes','unfreeze_panes'].includes(op.intent))return {ok:false,outcome:'xlsx-freeze-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
  if(typeof op.sheet!=='string'||!op.sheet.trim())return {ok:false,outcome:'xlsx-freeze-task-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
  const operation={index:0,intent:op.intent,sheet:op.sheet}
  if(op.intent==='freeze_panes'){
    if(!['rows','columns','at'].includes(op.mode))return {ok:false,outcome:'xlsx-freeze-task-invalid-mode',authority:'PLAN_ONLY',writeAllowed:false}
    operation.mode=op.mode
    if(op.mode==='at'){
      const p=parseA1Range(op.range)
      if(!p)return {ok:false,outcome:'xlsx-freeze-task-invalid-range',authority:'PLAN_ONLY',writeAllowed:false}
      operation.range=p.address
    }else{
      const max=op.mode==='rows'?1048576:16384
      if(!Number.isInteger(op.count)||op.count<1||op.count>max)return {ok:false,outcome:'xlsx-freeze-task-invalid-count',authority:'PLAN_ONLY',writeAllowed:false}
      operation.count=op.count
    }
  }
  return {ok:true,outcome:'xlsx-freeze-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function satisfied(x){return x?.ok===true&&x?.source==='live-coedit-editor'&&x?.verification?.measurable===true&&x.verification.match===true}
async function executeFreezeTask({task,api}){
  const plan=planTask(task);if(!plan.ok)return plan
  const op=plan.operation,initial=await api.inspect()
  if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-freeze-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
  const pre=await api.freezeObserved(op,false)
  if(!pre?.verification?.measurable)return {ok:false,outcome:'xlsx-freeze-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
  if(satisfied(pre)&&pre.noOp===true)return {ok:true,outcome:'xlsx-freeze-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:op,identityProof:'fresh-live-inventory+exact-freeze-bbox-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY'}}
  const fresh=await api.inspect()
  if(!fresh?.ok||fresh.authority!=='LIVE_READ'||!identity(fresh,op.sheet))return {ok:false,outcome:'xlsx-freeze-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
  const applied=await api.freezeObserved(op,true)
  if(!satisfied(applied)||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-freeze-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
  const finalInventory=await api.inspect()
  if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ'||!identity(finalInventory,op.sheet))return {ok:false,outcome:'xlsx-freeze-task-final-identity-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
  const final=await api.freezeObserved(op,false)
  if(!satisfied(final)||final.noOp!==true)return {ok:false,outcome:'xlsx-freeze-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
  return {ok:true,outcome:'xlsx-freeze-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:op,identityProof:'fresh-live-inventory+exact-freeze-bbox-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',verification:final.verification}}
}
module.exports={planTask,satisfied,executeFreezeTask}
