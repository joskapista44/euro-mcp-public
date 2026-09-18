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
 ['defined-name',['set_defined_name','rename_defined_name','delete_defined_name'],'DefinedName','definedNameObserved'],
 ['conditional-format',['add_conditional_format','delete_conditional_format'],'ConditionalFormat','cfObserved'],
 ['pivot',['create_pivot','refresh_pivot','delete_pivot_sheet'],'Pivot','pivotObserved'],
 ['chart',['set_chart','rename_chart','delete_chart'],'Chart','chartObserved']
]
const core={
 file:'core',
 intents:new Set(['create_sheet','write_range','copy_sheet','rename_sheet','delete_sheet']),
 agent:require('./xlsx-agent-task.cjs'),
 execute:'executeTask'
}
function clone(v){return JSON.parse(JSON.stringify(v))}
function creationFirstCoreOperations(operations){
 // A formula may reference another worksheet created by the same batch. The
 // editor resolves formula identity at write time, so every requested sheet
 // must exist before the first range write. Keep both partitions stable and
 // retain request indices for receipts and final verification.
 return [...operations.filter(op=>op.intent==='create_sheet'),...operations.filter(op=>op.intent!=='create_sheet')]
}
function coreFinalOperations(operations,tail){
 const final=clone(operations)
 for(const write of final.filter(op=>op.intent==='write_range')){
  const sheet=core.agent.finalSheetName(write.sheet,operations,write.index),wr=parseA1Range(write.range)
  for(const clear of tail.filter(op=>op?.intent==='clear_range'&&op.sheet===sheet)){
   const cr=parseA1Range(clear.range);if(!cr)continue
   const r0=Math.max(wr.start.row,cr.start.row),r1=Math.min(wr.end.row,cr.end.row),c0=Math.max(wr.start.column,cr.start.column),c1=Math.min(wr.end.column,cr.end.column)
   for(let r=r0;r<=r1;r++)for(let c=c0;c<=c1;c++){
    write.values[r-wr.start.row][c-wr.start.column]=null
    if(write.formulas)write.formulas[r-wr.start.row][c-wr.start.column]=null
   }
  }
 }
 return final
}
const families=definitions.map(([file,intents,fn,method])=>({
 file,intents,method,agent:require('./xlsx-agent-'+file+'-task.cjs'),
 transport:require('./xlsx-persistent-'+file+'.cjs'),execute:'execute'+fn+'Task'
}))
const {parseA1Range}=require('./range-reader.cjs')
function readbackPlan(task){
 if(task?.readbacks===undefined)return {ok:true,readbacks:[]}
 if(!Array.isArray(task.readbacks)||task.readbacks.length>10)return {ok:false,outcome:'xlsx-batch-invalid-readbacks',authority:'PLAN_ONLY'}
 const readbacks=[],seen=new Set()
 for(let i=0;i<task.readbacks.length;i++){
  const item=task.readbacks[i],parsed=parseA1Range(item?.range)
  if(typeof item?.sheet!=='string'||!item.sheet.trim()||!parsed||parsed.cellCount>400)return {ok:false,outcome:'xlsx-batch-invalid-readback',authority:'PLAN_ONLY',readbackIndex:i}
  const key=item.sheet.toLowerCase()+':'+parsed.address
  if(seen.has(key))return {ok:false,outcome:'xlsx-batch-duplicate-readback',authority:'PLAN_ONLY',readbackIndex:i}
  seen.add(key);readbacks.push({sheet:item.sheet,range:parsed.address})
 }
 return {ok:true,readbacks}
}
function goalTarget(family,operation){
 const sheet=String(operation.sheet||'').toLowerCase(),range=String(operation.range||'').toUpperCase()
 // Independent visual regions on one worksheet are compatible final-state
 // goals. Keeping the normalized range (and layout type) in their identity
 // permits a rich workbook to be composed in one editor session while exact
 // duplicate targets still fail closed.
 if(family.file==='format'||family.file==='merge')return family.file+':'+sheet+':'+range
 if(family.file==='layout')return family.file+':'+sheet+':'+range+':'+String(operation.type||'')
 if(family.file==='chart')return family.file+':'+sheet+':'+String(operation.name||operation.newName||'').toLowerCase()
 return family.file+':'+String(operation.sheet||operation.name).toLowerCase()
}
families.push(
 {file:'clear',intents:['clear_range'],agent:require('./xlsx-agent-clear-task.cjs'),execute:'executeClearTask',planTask:task=>{const op=task?.operations?.[0],p=op?parseA1Range(op.range):null;return task?.operations?.length===1&&op?.intent==='clear_range'&&typeof op.sheet==='string'&&op.sheet.trim()&&p?{ok:true,operation:{index:0,intent:'clear_range',sheet:op.sheet,range:p.address}}:{ok:false,outcome:'xlsx-clear-task-invalid',authority:'PLAN_ONLY'}}},
 {file:'move-sheet',intents:['move_sheet'],agent:require('./xlsx-agent-move-task.cjs'),execute:'executeMoveTask',planTask:task=>{const op=task?.operations?.[0],ok=task?.operations?.length===1&&op?.intent==='move_sheet'&&typeof op.sheet==='string'&&op.sheet.trim()&&typeof op.referenceSheet==='string'&&op.referenceSheet.trim()&&op.sheet!==op.referenceSheet&&['before','after'].includes(op.position);return ok?{ok:true,operation:{index:0,intent:'move_sheet',sheet:op.sheet,referenceSheet:op.referenceSheet,position:op.position}}:{ok:false,outcome:'xlsx-move-task-invalid',authority:'PLAN_ONLY'}}}
)
function planTask(task){
 if(!Array.isArray(task?.operations)||!task.operations.length||task.operations.length>100)return {ok:false,outcome:'xlsx-batch-invalid-operations',authority:'PLAN_ONLY'}
 const plannedReadbacks=readbackPlan(task);if(!plannedReadbacks.ok)return plannedReadbacks
 const steps=[],targets=new Set()
 const firstNonCore=task.operations.findIndex(op=>!core.intents.has(op?.intent))
 const coreCount=firstNonCore<0?task.operations.length:firstNonCore
 if(task.operations.slice(coreCount).some(op=>core.intents.has(op?.intent)))return {ok:false,outcome:'xlsx-batch-core-operations-must-come-first',authority:'PLAN_ONLY'}
 if(coreCount){
  const planned=core.agent.planTask({operations:task.operations.slice(0,coreCount)})
  if(!planned.ok)return planned
  const creationFirst=creationFirstCoreOperations(planned.operations)
  steps.push({index:0,family:'core',operations:creationFirst,verifyOperations:coreFinalOperations(creationFirst,task.operations.slice(coreCount))})
 }
 for(const [index,op] of task.operations.entries()){
  if(index<coreCount)continue
  const family=families.find(f=>f.intents.includes(op?.intent))
  if(!family)return {ok:false,outcome:'xlsx-batch-unsupported-intent',authority:'PLAN_ONLY',index}
  // AutoFit needs operation-bound retry tokens; not yet composed here.
  if(family.file==='layout'&&family.agent.isAutoFit(op))return {ok:false,outcome:'xlsx-batch-autofit-not-supported',authority:'PLAN_ONLY',index}
  const plan=(family.planTask||family.agent.planTask)({operations:[op]})
  if(!plan.ok)return {...plan,index}
  // Reject exact duplicate final-state targets while allowing independent ranges
  // and independently named objects to share a worksheet.
  const target=goalTarget(family,plan.operation)
  if(targets.has(target))return {ok:false,outcome:'xlsx-batch-conflicting-goals',authority:'PLAN_ONLY',index}
  targets.add(target);steps.push({index,family:family.file,operation:{...plan.operation,index}})
 }
 // Freeze panes are worksheet-view state in the deployed runtime. Pivot/chart
 // object work can switch the editor's sheet context after an otherwise
 // verified freeze. Apply and verify freeze goals last so the requested view
 // is the state persisted by the owning session.
 const orderedSteps=[...steps.filter(s=>s.family!=='freeze'),...steps.filter(s=>s.family==='freeze')]
 return {ok:true,outcome:'xlsx-batch-planned',authority:'PLAN_ONLY',steps:orderedSteps,readbacks:plannedReadbacks.readbacks}
}
function coreAdapter(api,readOnly){
 const blocked=async()=>({ok:false,outcome:'xlsx-batch-verification-write-blocked',authority:'PLAN_ONLY'})
 return {
  inspect:api.inspect,readRange:api.readRange,
  createSheetVerified:readOnly?blocked:api.createSheetVerified,
  writeRangeVerified:readOnly?blocked:api.writeRangeVerified,
  copySheetVerified:readOnly?blocked:api.copySheetVerified,
  renameSheetVerified:readOnly?blocked:api.renameSheetVerified,
  deleteSheetVerified:readOnly?blocked:api.deleteSheetVerified
 }
}
function adapter(api,family,readOnly){
 const blocked=async()=>({ok:false,outcome:'xlsx-batch-verification-write-blocked',authority:'PLAN_ONLY'})
 if(family.file==='clear')return {inspect:api.inspect,readRange:api.readRange,session:readOnly?undefined:api.session}
 if(family.file==='move-sheet')return {inspect:api.inspect,moveSheetVerified:readOnly?blocked:api.moveSheetVerified}
 const run=async(...args)=>{
  const spec=args[0],apply=family.file==='format'?!!spec.apply:!!args[1]
  if(family.file==='conditional-format'){
   const mutating=spec.type!=='cf.inspect'
   if(readOnly&&mutating)return {ok:false,outcome:'xlsx-batch-verification-write-blocked',source:'live-coedit-editor',authority:'PLAN_ONLY'}
   const r=await family.transport.runCommand(api.session,spec),out={...r,authority:family.transport.authorityFor(spec,r)}
   if(mutating&&out.ok&&out.authority==='LIVE_VERIFY')api.session.markWrite()
   return out
  }
  if(readOnly&&apply)return {ok:false,outcome:'xlsx-batch-verification-write-blocked',source:'live-coedit-editor'}
  const r=family.file==='format'
   ?await family.transport.runCommand(api.session,[spec.sheet,spec.range,spec.format,apply])
   :await family.transport.runCommand(api.session,spec,apply)
  if(apply&&(family.file==='chart'?(r?.applied||r?.mutationAttempted):(r?.ok&&!r.noOp)))api.session.markWrite()
  return r
 }
 // Do not expose session or mutators to the family executors.
 return {inspect:api.inspect,readRange:api.readRange,[family.method]:run}
}
async function executeBatchTask({task,api}){
 const plan=planTask(task);if(!plan.ok)return {...plan,writeAllowed:false}
 const steps=[],checks=[]
 for(const step of plan.steps){
  const f=step.family==='core'?core:families.find(f=>f.file===step.family)
  // Core writes are projected to the requested final state. A later clear of
  // the same cells makes the intermediate values dead writes on both apply
  // and retry, so they must never be dispatched.
  const operations=step.verifyOperations||step.operations||[step.operation]
  const result=await f.agent[f.execute]({task:{operations},api:step.family==='core'?coreAdapter(api,false):adapter(api,f,false)})
  steps.push({index:step.index,intent:step.family==='core'?'core_task':step.operation.intent,result})
  if(!result.ok||result.authority!=='LIVE_VERIFY')return {ok:false,outcome:'xlsx-batch-step-failed',authority:'LIVE_READ',writeAllowed:false,steps}
 }
 for(const step of plan.steps){
  const f=step.family==='core'?core:families.find(f=>f.file===step.family)
  const operations=step.verifyOperations||step.operations||[step.operation]
  const result=await f.agent[f.execute]({task:{operations},api:step.family==='core'?coreAdapter(api,true):adapter(api,f,true)})
  const ok=result.ok&&result.authority==='LIVE_VERIFY'&&result.noOp===true
  for(const op of operations)checks.push({index:op.index,intent:op.intent,ok})
  if(!ok)return {ok:false,outcome:'xlsx-batch-whole-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',writeAllowed:false,steps,checks,failed:result}
 }
 const readbacks=[]
 for(const spec of plan.readbacks){
  const observed=await api.readRange(spec)
  if(!observed?.ok||observed.authority!=='LIVE_READ')return {ok:false,outcome:'xlsx-batch-final-readback-failed',authority:'LIVE_READ',writeAllowed:false,steps,checks,failedReadback:spec}
  readbacks.push({authority:'LIVE_READ',source:observed.source||'live-coedit-editor',sheet:spec.sheet,range:spec.range,rows:observed.rows,columns:observed.columns,cells:observed.cells})
 }
 return {ok:true,outcome:'xlsx-batch-live-verified',authority:'LIVE_VERIFY',writeAllowed:false,noOp:steps.every(s=>s.result.noOp===true),steps,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',readOnly:true,checks,readbacks}}
}
async function executeBatchTaskInPersistentSession(options={}){
 const plan=planTask(options.task);if(!plan.ok)return {...plan,writeAllowed:false}
 return require('./xlsx-persistent-session.cjs').withPersistentXlsxSession(options,api=>executeBatchTask({task:options.task,api}))
}
module.exports={readbackPlan,goalTarget,planTask,creationFirstCoreOperations,coreFinalOperations,adapter,coreAdapter,executeBatchTask,executeBatchTaskInPersistentSession}
