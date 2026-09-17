'use strict'
const assert=require('assert')
const batch=require('./xlsx-persistent-batch.cjs')
const sort=require('./xlsx-persistent-sort.cjs')
const cf=require('./xlsx-persistent-conditional-format.cjs')
;(async()=>{
 const op={intent:'sort_range',sheet:'S',range:'A1:B3',keyRange:'A1:A3'}
 assert.equal(batch.planTask({operations:[op,{...op,order:'desc'}]}).outcome,'xlsx-batch-conflicting-goals')
 let reads=0
 const invalid=await batch.executeBatchTask({task:{operations:[op,{intent:'unknown'}]},api:{inspect(){reads++}}})
 assert.equal(invalid.ok,false);assert.equal(reads,0)
 assert.equal(batch.planTask({operations:[op,{intent:'create_sheet',name:'N'}]}).outcome,'xlsx-batch-core-operations-must-come-first')
 const coreTask={operations:[{intent:'create_sheet',name:'N'},{intent:'write_range',sheet:'N',range:'A1:B1',values:[['x',1]]}]}
 const sheets=new Set(['S']);let coreWrites=0,matrix=null
 const coreApi={
  inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[...sheets].map(name=>({name}))}),
  readRange:async()=>({ok:true,authority:'LIVE_READ',cells:matrix?matrix.map((row,r)=>row.map((value,c)=>({address:String.fromCharCode(65+c)+(r+1),rawValue:value,value,displayText:String(value)}))):[[{address:'A1',dataType:'blank'},{address:'B1',dataType:'blank'}]]}),
  createSheetVerified:async name=>{sheets.add(name);coreWrites++;return {ok:true,authority:'LIVE_VERIFY',noOp:false}},
  writeRangeVerified:async spec=>{matrix=spec.values;coreWrites++;return {ok:true,authority:'LIVE_VERIFY',noOp:false}},
  copySheetVerified:async()=>{throw Error('unexpected')},renameSheetVerified:async()=>{throw Error('unexpected')},deleteSheetVerified:async()=>{throw Error('unexpected')}
 }
 const coreFirst=await batch.executeBatchTask({task:coreTask,api:coreApi})
 assert.equal(coreFirst.ok,true);assert.equal(coreFirst.noOp,false);assert.equal(coreWrites,2);assert.equal(coreFirst.wholeTaskVerification.checks.length,2)
 const coreRetry=await batch.executeBatchTask({task:coreTask,api:coreApi})
 assert.equal(coreRetry.ok,true);assert.equal(coreRetry.noOp,true);assert.equal(coreWrites,2)
 const indexed=batch.planTask({operations:[...coreTask.operations,{intent:'format_range',sheet:'N',range:'A1:B1',format:{bold:true}},{intent:'set_defined_name',name:'N_R',refersTo:'=N!$A$1'}]})
 assert.deepEqual(indexed.steps.map(s=>s.operations?.map(x=>x.index)||s.operation.index),[[0,1],2,3])
 const overlaid=batch.planTask({operations:[{intent:'create_sheet',name:'X'},{intent:'write_range',sheet:'X',range:'A1:B2',values:[[1,2],[3,4]]},{intent:'rename_sheet',sheet:'X',name:'Y'},{intent:'clear_range',sheet:'Y',range:'B2'}]})
 assert.equal(overlaid.ok,true);assert.equal(overlaid.steps[0].operations[1].values[1][1],4);assert.equal(overlaid.steps[0].verifyOperations[1].values[1][1],null)
 const coreAgent=require('./xlsx-agent-task.cjs'),clearAgent=require('./xlsx-agent-clear-task.cjs'),originalCoreExecute=coreAgent.executeTask,originalClearExecute=clearAgent.executeClearTask,seen=[]
 coreAgent.executeTask=async({task})=>{seen.push(task.operations[1].values[1][1]);return {ok:true,authority:'LIVE_VERIFY',noOp:seen.length>1}}
 let clearCalls=0;clearAgent.executeClearTask=async()=>({ok:true,authority:'LIVE_VERIFY',noOp:++clearCalls>1})
 try{const overlayResult=await batch.executeBatchTask({task:{operations:[{intent:'create_sheet',name:'X'},{intent:'write_range',sheet:'X',range:'A1:B2',values:[[1,2],[3,4]]},{intent:'rename_sheet',sheet:'X',name:'Y'},{intent:'clear_range',sheet:'Y',range:'B2'}]},api:{}});assert.equal(overlayResult.ok,true);assert.deepEqual(seen,[4,null])}
 finally{coreAgent.executeTask=originalCoreExecute;clearAgent.executeClearTask=originalClearExecute}
 const originalCf=cf.runCommand;let cfPresent=false,cfWrites=0
 const cfRule={type:'xlCellValue',operator:'xlGreater',formula1:'5',fillColor:[1,2,3]}
 cf.runCommand=async(_session,spec)=>{
  if(spec.type==='cf.inspect')return {ok:true,rules:cfPresent?[{...cfRule,fillColor:66051,appliesTo:'A1:A3',index:0}]:[],count:cfPresent?1:0}
  cfPresent=true;return {ok:true,verification:{status:'PASS'}}
 }
 try{
  const cfApi={session:{markWrite(){cfWrites++}},inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}),readRange:async()=>({ok:true,authority:'LIVE_READ'})}
  const cfTask={operations:[{intent:'add_conditional_format',sheet:'S',range:'A1:A3',rule:cfRule}]}
  const cfFirst=await batch.executeBatchTask({task:cfTask,api:cfApi});assert.equal(cfFirst.ok,true);assert.equal(cfFirst.noOp,false);assert.equal(cfWrites,1)
  const cfRetry=await batch.executeBatchTask({task:cfTask,api:cfApi});assert.equal(cfRetry.ok,true);assert.equal(cfRetry.noOp,true);assert.equal(cfWrites,1)
 }finally{cf.runCommand=originalCf}
 const original=sort.runCommand
 let ordered=false,writes=0,applyCalls=0
 const session={markWrite(){writes++}}
 const api={session,inspect:async()=>({ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]})}
 sort.runCommand=async(_session,_spec,apply)=>{
  if(apply){applyCalls++;ordered=true}
  return {ok:true,source:'live-coedit-editor',noOp:!apply&&ordered,applied:apply,state:{ordered},verification:{measurable:true,match:ordered,keys:["A","B"]}}
 }
 try{
  const first=await batch.executeBatchTask({task:{operations:[op]},api})
  assert.equal(first.ok,true);assert.equal(first.noOp,false);assert.equal(writes,1);assert.equal(first.wholeTaskVerification.readOnly,true)
  const retry=await batch.executeBatchTask({task:{operations:[op]},api})
  assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(writes,1)
  // External state changes only when the final read-only verification starts.
  let observations=0
  sort.runCommand=async(_s,_op,apply)=>{
   if(apply){applyCalls++;throw new Error('final verifier must not write')}
   observations++
   const match=observations===1
   return {ok:true,source:'live-coedit-editor',noOp:match,state:{ordered:false},verification:{measurable:true,match,keys:["A","B"]}}
  }
  const failed=await batch.executeBatchTask({task:{operations:[op]},api})
  assert.equal(failed.outcome,'xlsx-batch-whole-verify-failed');assert.equal(applyCalls,1)
 }finally{sort.runCommand=original}
 console.log('XLSX BATCH STATIC: PASS (preflight, core create/write, conflicts, retry, read-only final verification)')
})().catch(e=>{console.error(e);process.exitCode=1})
