'use strict'
const {parseA1Range}=require('./range-reader.cjs')
const NAME=/^[A-Za-z_][A-Za-z0-9_.]*$/
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||!['create_pivot','delete_pivot_sheet'].includes(op.intent))return {ok:false,outcome:'xlsx-pivot-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.name!=='string'||!NAME.test(op.name))return {ok:false,outcome:'xlsx-pivot-task-invalid-name',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='delete_pivot_sheet'){
  if(typeof op.pivotSheet!=='string'||!op.pivotSheet.trim())return {ok:false,outcome:'xlsx-pivot-task-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
  return {ok:true,outcome:'xlsx-pivot-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,name:op.name,pivotSheet:op.pivotSheet}}
 }
 if(typeof op.sourceSheet!=='string'||!op.sourceSheet.trim())return {ok:false,outcome:'xlsx-pivot-task-source-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
 const range=parseA1Range(op.sourceRange);if(!range)return {ok:false,outcome:'xlsx-pivot-task-invalid-source-range',authority:'PLAN_ONLY',writeAllowed:false}
 for(const k of ['rowField','columnField','dataField','styleName'])if(typeof op[k]!=='string'||!op[k])return {ok:false,outcome:'xlsx-pivot-task-field-required',authority:'PLAN_ONLY',writeAllowed:false,field:k}
 if(!Array.isArray(op.assertions)||op.assertions.length<1||op.assertions.some(a=>!Array.isArray(a?.items)||!a.items.length||a.expected==null))return {ok:false,outcome:'xlsx-pivot-task-semantic-assertions-required',authority:'PLAN_ONLY',writeAllowed:false}
 return {ok:true,outcome:'xlsx-pivot-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,name:op.name,sourceSheet:op.sourceSheet,sourceRange:range.address,rowField:op.rowField,columnField:op.columnField,dataField:op.dataField,styleName:op.styleName,assertions:op.assertions.map(a=>({items:a.items.map(String),expected:String(a.expected)}))}}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executePivotTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||(op.intent==='create_pivot'&&!identity(initial,op.sourceSheet)))return {ok:false,outcome:'xlsx-pivot-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.pivotObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-pivot-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-pivot-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{name:op.name,pivotSheet:op.pivotSheet},identityProof:'fresh-live-inventory+exact-pivot-semantic-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ'||(op.intent==='create_pivot'&&!identity(freshInventory,op.sourceSheet)))return {ok:false,outcome:'xlsx-pivot-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.pivotObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.state)!==JSON.stringify(pre.state))return {ok:false,outcome:'xlsx-pivot-task-state-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.pivotObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-pivot-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-pivot-task-final-read-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.pivotObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-pivot-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-pivot-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{name:op.name,pivotSheet:final.state?.parentSheet||op.pivotSheet},identityProof:'fresh-pivot-fingerprint+public-GetData-semantic-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={NAME,planTask,identity,measured,executePivotTask}
