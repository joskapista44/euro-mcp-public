'use strict'
const assert=require('assert/strict')
const batch=require('./xlsx-persistent-batch.cjs')
const core=require('./xlsx-agent-task.cjs'),move=require('./xlsx-agent-range-move-task.cjs')
const retry=require('./xlsx-range-move-retry-token.cjs')
async function main(){
 const op={intent:'move_range',sheet:'D',range:'A1:B1',targetSheet:'D',targetRange:'C1:D1'},fp='0'.repeat(64)
 const token={version:2,operationFingerprint:retry.operationFingerprint(op),sourceBeforeFingerprint:fp,targetBeforeFingerprint:fp,sourceAfterFingerprint:fp,targetAfterFingerprint:fp}
 const originalCore=core.executeTask,originalMove=move.executeRangeMoveTask
 let reads=0
 core.executeTask=async({task})=>{
  reads++
  assert.deepEqual(task.operations[1].values,[[null,null]],'reopened retry must not restore consumed source cells')
  return {ok:true,authority:'LIVE_VERIFY',noOp:true}
 }
 move.executeRangeMoveTask=async()=>({ok:true,authority:'LIVE_VERIFY',noOp:true,retryToken:token})
 try{
  const result=await batch.executeBatchTask({task:{operations:[{intent:'create_sheet',name:'D'},{intent:'write_range',sheet:'D',range:'A1:B1',values:[['source',1]]},{...op,retryToken:token}]},api:{}})
  assert.equal(result.ok,true);assert.equal(result.noOp,true);assert.equal(reads,2)
 }finally{core.executeTask=originalCore;move.executeRangeMoveTask=originalMove}
 console.log('XLSX BATCH OBJECT RECEIPT STATIC: PASS')
}
main().catch(error=>{console.error(error);process.exitCode=1})
