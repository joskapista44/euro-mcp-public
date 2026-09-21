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
 ['chart',['set_chart','rename_chart','delete_chart'],'Chart','chartObserved'],
 ['page-layout',['set_page_layout'],'PageLayout','pageLayoutObserved'],
 ['print-setup',['set_print_setup'],'PrintSetup','printSetupObserved'],
 ['print-titles',['set_print_titles'],'PrintTitles','printTitlesObserved'],
 ['print-area',['set_print_area'],'PrintArea','printAreaObserved'],
 ['header-footer',['set_header_footer'],'HeaderFooter','headerFooterObserved'],
 ['page-break',['set_page_break'],'PageBreak','pageBreakObserved'],
 ['first-page-number',['set_first_page_number'],'FirstPageNumber','firstPageNumberObserved']
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
function rangeToA1(r){const col=n=>{let s='';for(let x=n+1;x;x=Math.floor((x-1)/26))s=String.fromCharCode(65+(x-1)%26)+s;return s};return col(r.start.column)+(r.start.row+1)+':'+col(r.end.column)+(r.end.row+1)}
function projectCoreWritesForStructural(final,tail){
 let ops=final
 for(const st of tail.filter(op=>['insert_rows','insert_columns'].includes(op?.intent))){
  const sr=parseA1Range(st.range);if(!sr)continue
  const next=[]
  for(const op of ops){
   if(op.intent!=='write_range'||op.sheet!==st.sheet){next.push(op);continue}
   const wr=parseA1Range(op.range);if(!wr){next.push(op);continue}
   if(st.intent==='insert_rows'){
    const count=sr.end.row-sr.start.row+1
    if(sr.start.row<=wr.start.row){const x=clone(op);wr.start.row+=count;wr.end.row+=count;x.range=rangeToA1(wr);next.push(x)}
    else if(sr.start.row<=wr.end.row){const cut=sr.start.row-wr.start.row,top=clone(op),bottom=clone(op);top.range=rangeToA1({start:{...wr.start},end:{row:sr.start.row-1,column:wr.end.column}});top.values=op.values.slice(0,cut);if(top.formulas)top.formulas=op.formulas.slice(0,cut);bottom.range=rangeToA1({start:{row:sr.start.row+count,column:wr.start.column},end:{row:wr.end.row+count,column:wr.end.column}});bottom.values=op.values.slice(cut);if(bottom.formulas)bottom.formulas=op.formulas.slice(cut);next.push(top,bottom)}
    else next.push(op)
   }else{
    const count=sr.end.column-sr.start.column+1
    if(sr.start.column<=wr.start.column){const x=clone(op);wr.start.column+=count;wr.end.column+=count;x.range=rangeToA1(wr);next.push(x)}
    else if(sr.start.column<=wr.end.column){const cut=sr.start.column-wr.start.column,left=clone(op),right=clone(op);left.range=rangeToA1({start:{...wr.start},end:{row:wr.end.row,column:sr.start.column-1}});left.values=op.values.map(r=>r.slice(0,cut));if(left.formulas)left.formulas=op.formulas.map(r=>r.slice(0,cut));right.range=rangeToA1({start:{row:wr.start.row,column:sr.start.column+count},end:{row:wr.end.row,column:wr.end.column+count}});right.values=op.values.map(r=>r.slice(cut));if(right.formulas)right.formulas=op.formulas.map(r=>r.slice(cut));next.push(left,right)}
    else next.push(op)
   }
  }
  ops=next
 }
 return ops
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
  // A later range move consumes the source cells. Project that destructive
  // final state into earlier core writes so whole-task verification checks the
  // requested final workbook rather than the pre-move intermediate state.
  for(const move of tail.filter(op=>op?.intent==='move_range'&&op.sheet===sheet)){
   const mr=parseA1Range(move.range);if(!mr)continue
   const r0=Math.max(wr.start.row,mr.start.row),r1=Math.min(wr.end.row,mr.end.row),c0=Math.max(wr.start.column,mr.start.column),c1=Math.min(wr.end.column,mr.end.column)
   if(r0>r1||c0>c1)continue
   for(let r=r0;r<=r1;r++)for(let c=c0;c<=c1;c++){write.values[r-wr.start.row][c-wr.start.column]=null;if(write.formulas)write.formulas[r-wr.start.row][c-wr.start.column]=null}
  }
  // A later sort changes the final row order of an earlier core write. Do not
  // verify/retry against the pre-sort matrix. Project the deterministic sort
  // into the core write's expected final state when the sort range is fully
  // contained in that write and uses a single-column key.
  for(const sort of tail.filter(op=>op?.intent==='sort_range'&&op.sheet===sheet)){
   const sr=parseA1Range(sort.range),kr=parseA1Range(sort.keyRange);if(!sr||!kr||sr.start.row<wr.start.row||sr.end.row>wr.end.row||sr.start.column<wr.start.column||sr.end.column>wr.end.column)continue
   const from=sr.start.row-wr.start.row,to=sr.end.row-wr.start.row,key=kr.start.column-wr.start.column,header=sort.hasHeaders!==false?1:0
   const start=from+header,rows=[]
   for(let r=start;r<=to;r++)rows.push({values:write.values[r],formulas:write.formulas?write.formulas[r]:null})
   const cell=x=>{const f=x.formulas?.[key];return f!=null?f:x.values?.[key]}
   rows.sort((a,b)=>{const av=cell(a),bv=cell(b);if(av==null&&bv==null)return 0;if(av==null)return 1;if(bv==null)return -1;const an=Number(av),bn=Number(bv),cmp=Number.isFinite(an)&&Number.isFinite(bn)?an-bn:String(av).localeCompare(String(bv));return sort.order==='desc'?-cmp:cmp})
   rows.forEach((row,i)=>{write.values[start+i]=row.values;if(write.formulas)write.formulas[start+i]=row.formulas})
  }
 }
 return projectCoreWritesForStructural(final,tail)
}
const families=definitions.map(([file,intents,fn,method])=>({
 file,intents,method,agent:require('./xlsx-agent-'+file+'-task.cjs'),
 transport:require(file==='first-page-number'?'./xlsx-persistent-page-number-order.cjs':'./xlsx-persistent-'+file+'.cjs'),execute:'execute'+fn+'Task'
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
 if(family.file==='conditional-format'){
  const r=operation.rule||{},value=v=>v===undefined?'*':String(v==null?'':v),fill=r.fillColor===undefined?'*':Array.isArray(r.fillColor)?r.fillColor.join(','):String(r.fillColor)
  return family.file+':'+sheet+':'+range+':'+[value(r.type),value(r.operator),value(r.formula1),value(r.formula2),value(r.priority),fill].join(':')
 }
 if(family.file==='chart')return family.file+':'+sheet+':'+String(operation.name||operation.newName||'').toLowerCase()
 if(family.file==='page-layout')return family.file+':'+sheet
 if(family.file==='print-setup')return family.file+':'+sheet+':'+String(operation.mode||'')
 if(family.file==='print-titles')return family.file+':'+sheet+':'+String(operation.axis||'')
 if(family.file==='print-area')return family.file+':'+sheet
 if(family.file==='header-footer')return family.file+':'+sheet+':'+String(operation.slot||'')
 if(family.file==='page-break')return family.file+':'+sheet+':'+String(operation.mode||'')+':'+String(operation.axis||'')+':'+String(operation.at||'')
 if(family.file==='first-page-number')return family.file+':'+sheet
 return family.file+':'+String(operation.sheet||operation.name).toLowerCase()
}
families.push(
 {file:'clear',intents:['clear_range'],agent:require('./xlsx-agent-clear-task.cjs'),execute:'executeClearTask',planTask:task=>{const op=task?.operations?.[0],p=op?parseA1Range(op.range):null;return task?.operations?.length===1&&op?.intent==='clear_range'&&typeof op.sheet==='string'&&op.sheet.trim()&&p?{ok:true,operation:{index:0,intent:'clear_range',sheet:op.sheet,range:p.address}}:{ok:false,outcome:'xlsx-clear-task-invalid',authority:'PLAN_ONLY'}}},
 {file:'range-copy',intents:['copy_range'],agent:require('./xlsx-agent-range-copy-task.cjs'),execute:'executeRangeCopyTask'},
 {file:'range-move',intents:['move_range'],agent:require('./xlsx-agent-range-move-task.cjs'),execute:'executeRangeMoveTask'},
 {file:'structural',intents:['insert_rows','delete_rows','insert_columns','delete_columns'],agent:require('./xlsx-agent-structural-task.cjs'),execute:'executeStructuralTaskInSession',persistent:require('./xlsx-persistent-structural.cjs')},
 {file:'move-sheet',intents:['move_sheet'],agent:require('./xlsx-agent-move-task.cjs'),execute:'executeMoveTask',planTask:task=>{const op=task?.operations?.[0],ok=task?.operations?.length===1&&op?.intent==='move_sheet'&&typeof op.sheet==='string'&&op.sheet.trim()&&typeof op.referenceSheet==='string'&&op.referenceSheet.trim()&&op.sheet!==op.referenceSheet&&['before','after'].includes(op.position);return ok?{ok:true,operation:{index:0,intent:'move_sheet',sheet:op.sheet,referenceSheet:op.referenceSheet,position:op.position}}:{ok:false,outcome:'xlsx-move-task-invalid',authority:'PLAN_ONLY'}}}
)
function planTask(task){
 if(!Array.isArray(task?.operations)||!task.operations.length||task.operations.length>1000)return {ok:false,outcome:'xlsx-batch-invalid-operations',authority:'PLAN_ONLY'}
 const plannedReadbacks=readbackPlan(task);if(!plannedReadbacks.ok)return plannedReadbacks
 const steps=[],targets=new Set()
 const firstNonCore=task.operations.findIndex(op=>!core.intents.has(op?.intent))
 const coreCount=firstNonCore<0?task.operations.length:firstNonCore
 if(task.operations.slice(coreCount).some(op=>core.intents.has(op?.intent)))return {ok:false,outcome:'xlsx-batch-core-operations-must-come-first',authority:'PLAN_ONLY'}
 if(coreCount){
  const planned=core.agent.planTask({operations:task.operations.slice(0,coreCount)})
  if(!planned.ok)return planned
  const creationFirst=creationFirstCoreOperations(planned.operations)
  const finalCore=coreFinalOperations(creationFirst,task.operations.slice(coreCount))
  // Destructive later operations (currently range_move) change what an earlier
  // core write means on retry. Dispatch the projected final-state core task as
  // well as verifying against it; otherwise retry resurrects the consumed source
  // and forces the move primitive to write again.
  steps.push({index:0,family:'core',operations:finalCore,verifyOperations:finalCore})
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
 if(family.file==='range-copy')return {inspect:api.inspect,readRange:api.readRange,copyRangeVerified:readOnly?blocked:api.copyRangeVerified}
 if(family.file==='range-move')return {inspect:api.inspect,readRange:api.readRange,moveRangeVerified:readOnly?blocked:api.moveRangeVerified}
 if(family.file==='structural')return {inspect:api.inspect,readRange:api.readRange,session:readOnly?undefined:api.session}
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
  let operations=step.verifyOperations||step.operations||[step.operation]
  // Structural standalone tasks intentionally accept a missing precondition and
  // derive the before/post proof from the owning LIVE session. Batch operations
  // are normalized by planTask, where absent optional fingerprints are null;
  // strip those null optionals before handing the operation back to the strict
  // structural task planner (undefined = omitted, null = invalid supplied value).
  if(step.family==='structural')operations=operations.map(op=>{const x={...op};if(x.preconditionFingerprint==null)delete x.preconditionFingerprint;if(x.expectedPostFingerprint==null)delete x.expectedPostFingerprint;if(x.retryToken==null)delete x.retryToken;return x})
  // Range-move retry idempotence is handled by the move family itself. The
  // batch must not synthesize a retry token from only the current post-state;
  // that cannot reconstruct the operation-bound before/after proof.
  const result=step.family==='structural'
   ?await f.persistent[f.execute](api.session,api,{operations})
   :await f.agent[f.execute]({task:{operations},api:step.family==='core'?coreAdapter(api,false):adapter(api,f,false)})
  steps.push({index:step.index,intent:step.family==='core'?'core_task':step.operation.intent,result})
  if(!result.ok||result.authority!=='LIVE_VERIFY')return {ok:false,outcome:'xlsx-batch-step-failed',authority:'LIVE_READ',writeAllowed:false,steps}
 }
 for(const step of plan.steps){
  const f=step.family==='core'?core:families.find(f=>f.file===step.family)
  const operations=step.verifyOperations||step.operations||[step.operation]
  let verifyOperations=operations
  if(step.family==='range-move'){const applied=steps.find(s=>s.index===step.index&&s.intent==='move_range')?.result;if(applied?.retryToken)verifyOperations=[{...operations[0],retryToken:applied.retryToken}]}
  if(step.family==='structural'){const applied=steps.find(s=>s.index===step.index&&s.intent===step.operation.intent)?.result;if(applied?.retryToken)verifyOperations=[{...operations[0],retryToken:applied.retryToken}];verifyOperations=verifyOperations.map(op=>{const x={...op};if(x.preconditionFingerprint==null)delete x.preconditionFingerprint;if(x.expectedPostFingerprint==null)delete x.expectedPostFingerprint;if(x.retryToken==null)delete x.retryToken;return x})}
  const result=step.family==='structural'
   ?await f.persistent[f.execute](api.session,api,{operations:verifyOperations})
   :await f.agent[f.execute]({task:{operations:verifyOperations},api:step.family==='core'?coreAdapter(api,true):adapter(api,f,true)})
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
