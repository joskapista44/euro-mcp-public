'use strict'
// Composes final-state intents in ONE already-open session. The owning wrapper
// alone saves and closes. Verification adapters cannot dispatch mutations.
const definitions=[
 ['format',['format_range'],'Format','formatRangeObserved'],
 ['layout',['layout_range'],'Layout','layoutObserved'],
 ['merge',['merge_range','unmerge_range'],'Merge','mergeObserved'],
 ['freeze',['freeze_panes','unfreeze_panes'],'Freeze','freezeObserved'],
 ['sort',['sort_range'],'Sort','sortObserved'],
 ['filter',['filter_range','clear_filter'],'Filter','filterObserved'],
 ['validation',['set_validation','clear_validation'],'Validation','validationObserved'],
 ['defined-name',['set_defined_name','delete_defined_name'],'DefinedName','definedNameObserved']
]
const families=definitions.map(([file,intents,fn,method])=>({
 file,intents,method,agent:require('./xlsx-agent-'+file+'-task.cjs'),
 transport:require('./xlsx-persistent-'+file+'.cjs'),execute:'execute'+fn+'Task'
}))
function planTask(task){
 if(!Array.isArray(task?.operations)||!task.operations.length||task.operations.length>100)return {ok:false,outcome:'xlsx-batch-invalid-operations',authority:'PLAN_ONLY'}
 const steps=[],targets=new Set()
 for(const [index,op] of task.operations.entries()){
  const family=families.find(f=>f.intents.includes(op?.intent))
  if(!family)return {ok:false,outcome:'xlsx-batch-unsupported-intent',authority:'PLAN_ONLY',index}
  // AutoFit needs operation-bound retry tokens; not yet composed here.
  if(family.file==='layout'&&family.agent.isAutoFit(op))return {ok:false,outcome:'xlsx-batch-autofit-not-supported',authority:'PLAN_ONLY',index}
  const plan=family.agent.planTask({operations:[op]})
  if(!plan.ok)return {...plan,index}
  // A conservative initial contract: one final-state goal per family per sheet,
  // or per workbook name. Avoid replaying overwritten intermediate goals.
  const target=family.file+':'+String(op.sheet||op.name).toLowerCase()
  if(targets.has(target))return {ok:false,outcome:'xlsx-batch-conflicting-goals',authority:'PLAN_ONLY',index}
  targets.add(target);steps.push({index,family:family.file,operation:plan.operation})
 }
 return {ok:true,outcome:'xlsx-batch-planned',authority:'PLAN_ONLY',steps}
}
function adapter(api,family,readOnly){
 const run=async(...args)=>{
  const spec=args[0],apply=family.file==='format'?!!spec.apply:!!args[1]
  if(readOnly&&apply)return {ok:false,outcome:'xlsx-batch-verification-write-blocked',source:'live-coedit-editor'}
  const r=family.file==='format'
   ?await family.transport.runCommand(api.session,[spec.sheet,spec.range,spec.format,apply])
   :await family.transport.runCommand(api.session,spec,apply)
  if(apply&&r?.ok&&!r.noOp)api.session.markWrite()
  return r
 }
 // Do not expose session or mutators to the family executors.
 return {inspect:api.inspect,readRange:api.readRange,[family.method]:run}
}
async function executeBatchTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return {...plan,writeAllowed:false}
 const steps=[],checks=[]
 for(const step of plan.steps){
  const f=families.find(f=>f.file===step.family)
  const result=await f.agent[f.execute]({task:{operations:[step.operation]},api:adapter(api,f,false)})
  steps.push({index:step.index,intent:step.operation.intent,result})
  if(!result.ok||result.authority!=='LIVE_VERIFY')return {ok:false,outcome:'xlsx-batch-step-failed',authority:'LIVE_READ',writeAllowed:false,steps}
 }
 for(const step of plan.steps){
  const f=families.find(f=>f.file===step.family)
  const result=await f.agent[f.execute]({task:{operations:[step.operation]},api:adapter(api,f,true)})
  checks.push({index:step.index,intent:step.operation.intent,ok:result.ok&&result.authority==='LIVE_VERIFY'&&result.noOp===true})
  if(!checks.at(-1).ok)return {ok:false,outcome:'xlsx-batch-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,steps,checks,failed:result}
 }
 return {ok:true,outcome:'xlsx-batch-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:steps.every(s=>s.result.noOp===true),steps,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',readOnly:true,checks}}
}
async function executeBatchTaskInPersistentSession(options={}){
 const plan=planTask(options.task);if(!plan.ok)return {...plan,writeAllowed:false}
 return require('./xlsx-persistent-session.cjs').withPersistentXlsxSession(options,api=>executeBatchTask({task:options.task,api}))
}
module.exports={planTask,adapter,executeBatchTask,executeBatchTaskInPersistentSession}
