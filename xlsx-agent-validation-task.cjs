'use strict'
const {parseA1Range}=require('./range-reader.cjs')
const VALIDATION_TYPES=new Set(['xlValidateCustom','xlValidateDate','xlValidateDecimal','xlValidateList','xlValidateTextLength','xlValidateTime','xlValidateWholeNumber'])
const ALERT_STYLES=new Set(['xlValidAlertInformation','xlValidAlertStop','xlValidAlertWarning'])
const OPERATORS=new Set(['xlBetween','xlEqual','xlGreater','xlGreaterEqual','xlLess','xlLessEqual','xlNotBetween','xlNotEqual'])
function planTask(task){
 const op=task?.operations?.length===1?task.operations[0]:null
 if(!op||!['set_validation','clear_validation'].includes(op.intent))return {ok:false,outcome:'xlsx-validation-task-single-operation-required',authority:'PLAN_ONLY',writeAllowed:false}
 if(typeof op.sheet!=='string'||!op.sheet.trim())return {ok:false,outcome:'xlsx-validation-task-sheet-required',authority:'PLAN_ONLY',writeAllowed:false}
 const p=parseA1Range(op.range);if(!p)return {ok:false,outcome:'xlsx-validation-task-invalid-range',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.intent==='clear_validation')return {ok:true,outcome:'xlsx-validation-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation:{index:0,intent:op.intent,sheet:op.sheet,range:p.address}}
 if(!VALIDATION_TYPES.has(op.validationType))return {ok:false,outcome:'xlsx-validation-task-invalid-type',authority:'PLAN_ONLY',writeAllowed:false}
 const alertStyle=op.alertStyle||'xlValidAlertStop',operator=op.operator||'xlBetween'
 if(!ALERT_STYLES.has(alertStyle))return {ok:false,outcome:'xlsx-validation-task-invalid-alert-style',authority:'PLAN_ONLY',writeAllowed:false}
 if(!OPERATORS.has(operator))return {ok:false,outcome:'xlsx-validation-task-invalid-operator',authority:'PLAN_ONLY',writeAllowed:false}
 if(op.formula1==null)return {ok:false,outcome:'xlsx-validation-task-formula-required',authority:'PLAN_ONLY',writeAllowed:false}
 const unsupported=['ignoreBlank','inCellDropdown','inputMessage','inputTitle','showError','showInput','errorMessage','errorTitle']
 for(const option of unsupported)if(Object.hasOwn(op,option))return {ok:false,outcome:'xlsx-validation-task-option-not-persistable',authority:'PLAN_ONLY',writeAllowed:false,option}
 const str=v=>v==null?'':String(v)
 const operation={index:0,intent:op.intent,sheet:op.sheet,range:p.address,validationType:op.validationType,alertStyle,operator,formula1:str(op.formula1),formula2:str(op.formula2)}
 return {ok:true,outcome:'xlsx-validation-task-planned',authority:'PLAN_ONLY',writeAllowed:true,operation}
}
function identity(inv,name){return (inv?.sheets||[]).filter(x=>x?.name===name).length===1}
function measured(r){return r?.ok===true&&r?.source==='live-coedit-editor'&&r?.verification?.measurable===true}
async function executeValidationTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return plan
 const op=plan.operation,initial=await api.inspect()
 if(!initial?.ok||initial.authority!=='LIVE_READ'||!identity(initial,op.sheet))return {ok:false,outcome:'xlsx-validation-task-initial-identity-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const pre=await api.validationObserved(op,false)
 if(!measured(pre))return {ok:false,outcome:pre?.outcome||'xlsx-validation-task-precheck-unverifiable',authority:'LIVE_READ',writeAllowed:false,plan,pre}
 if(pre.verification.match&&pre.noOp===true)return {ok:true,outcome:'xlsx-validation-task-already-satisfied',authority:'LIVE_VERIFY',writeAllowed:false,noOp:true,plan,receipt:[{index:0,intent:op.intent,status:'ALREADY_SATISFIED',resolvedTarget:{sheet:op.sheet,range:op.range},identityProof:'fresh-live-inventory+exact-validation-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}}
 const freshInventory=await api.inspect()
 if(!freshInventory?.ok||freshInventory.authority!=='LIVE_READ'||!identity(freshInventory,op.sheet))return {ok:false,outcome:'xlsx-validation-task-fresh-boundary-failed',authority:'PLAN_ONLY',writeAllowed:false,plan}
 const fresh=await api.validationObserved(op,false)
 if(!measured(fresh)||JSON.stringify(fresh.state)!==JSON.stringify(pre.state))return {ok:false,outcome:'xlsx-validation-task-state-changed-before-mutation',authority:'LIVE_READ',writeAllowed:false,plan,pre,fresh}
 const applied=await api.validationObserved(op,true)
 if(!measured(applied)||!applied.verification.match||applied.noOp===true||applied.applied!==true)return {ok:false,outcome:'xlsx-validation-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,plan,applied}
 const finalInventory=await api.inspect()
 if(!finalInventory?.ok||finalInventory.authority!=='LIVE_READ'||!identity(finalInventory,op.sheet))return {ok:false,outcome:'xlsx-validation-task-final-identity-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan}
 const final=await api.validationObserved(op,false)
 if(!measured(final)||!final.verification.match||final.noOp!==true)return {ok:false,outcome:'xlsx-validation-task-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,plan,final}
 return {ok:true,outcome:'xlsx-validation-task-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:false,plan,receipt:[{index:0,intent:op.intent,status:'APPLIED',resolvedTarget:{sheet:op.sheet,range:op.range},identityProof:'fresh-live-inventory+validation-state-fingerprint+exact-validation-readback',verification:'LIVE_VERIFY'}],wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:final.state}}
}
module.exports={VALIDATION_TYPES,ALERT_STYLES,OPERATORS,planTask,identity,measured,executeValidationTask}
