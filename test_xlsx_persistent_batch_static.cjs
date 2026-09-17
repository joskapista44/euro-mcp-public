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
 console.log('XLSX BATCH STATIC: PASS (preflight, conflicts, retry, read-only final verification)')
})().catch(e=>{console.error(e);process.exitCode=1})
