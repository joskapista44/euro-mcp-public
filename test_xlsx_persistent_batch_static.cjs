'use strict'
const assert=require('assert')
const batch=require('./xlsx-persistent-batch.cjs')
const sort=require('./xlsx-persistent-sort.cjs')
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
