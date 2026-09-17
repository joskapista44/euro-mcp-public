'use strict'
const {parseA1Range}=require('./range-reader.cjs')
const OPERATORS=new Set(['xlAnd','xlOr','xlBottom10Items','xlBottom10Percent','xlFilterCellColor','xlFilterDynamic','xlFilterFontColor','xlFilterIcon','xlFilterValues','xlTop10Items','xlTop10Percent'])
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||op.intent!=='filter_range')return {ok:false,outcome:'xlsx-filter-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.sheet!=='string'||!op.sheet.trim())return {ok:false,outcome:'xlsx-filter-task-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
 const p=parseA1Range(op.range);if(!p)return {ok:false,outcome:'xlsx-filter-task-invalid-range',authority:'PLAN_ONLY',writeAllowed:false}
 const width=p.end.column-p.start.column+1
 if(!Number.isInteger(op.field)||op.field<1||op.field>width)return {ok:false,outcome:'xlsx-filter-task-invalid-field',authority:'PLAN_ONLY',writeAllowed:false}
 const operator=op.operator||'xlOr';if(!OPERATORS.has(operator))return {ok:false,outcome:'xlsx-filter-task-invalid-operator',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.criteria1==null)return {ok:false,outcome:'xlsx-filter-task-criteria-required',authority:'PLAN_ONLY',writeAllowed:false}
 return {ok:true,outcome:'xlsx-filter-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:'filter_range',sheet:op.sheet,range:p.address,field:op.field,criteria1:op.criteria1,criteria2:op.criteria2??null,operator}}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeFilterTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-filter-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.filterObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-filter-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-filter-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{sheet:op.sheet,range:op.range,field:op.field},identityProof:'fresh-live-inventory+exact-filter-model-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ'||!identity(freshInventory,op.sheet))return {ok:false,outcome:'xlsx-filter-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.filterObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.state)!==JSON.stringify(pre.state))return {ok:false,outcome:'xlsx-filter-task-state-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.filterObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-filter-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ'||!identity(finalInventory,op.sheet))return {ok:false,outcome:'xlsx-filter-task-final-identity-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.filterObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-filter-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-filter-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,range:op.range,field:op.field},identityProof:'fresh-live-inventory+filter-state-fingerprint+exact-filter-model-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={OPERATORS,planTask,identity,measured,executeFilterTask}
