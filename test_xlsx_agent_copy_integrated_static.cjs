'use strict'
const assert=require('assert')
const task=require('./xlsx-agent-task.cjs')
function inv(names){return {ok:true,authority:'LIVE_READ',sheets:names.map((name,index)=>({name,index,usedRange:'A1'}))}}
function read(v){return {ok:true,authority:'LIVE_READ',rows:1,columns:1,cells:[[{rawValue:v,value:v,displayText:String(v),formula:null,dataType:typeof v}]]}}
async function main(){
  const planned=task.planTask({operations:[{intent:'copy_sheet',sheet:'Source',name:'Copy'},{intent:'write_range',sheet:'Copy',range:'A1',values:[[9]]},{intent:'rename_sheet',sheet:'Copy',name:'Final'}]})
  assert.equal(planned.ok,true)
  assert.deepEqual(planned.operations.map(x=>x.intent),['copy_sheet','write_range','rename_sheet'])
  assert.equal(task.planTask({operations:[{intent:'copy_sheet',sheet:'A',name:'B'},{intent:'delete_sheet',sheet:'B'}]}).outcome,'xlsx-task-copy-then-delete-not-supported')
  assert.equal(task.planTask({operations:[{intent:'copy_sheet',sheet:'A',name:'B'},{intent:'copy_sheet',sheet:'B',name:'C'}]}).outcome,'xlsx-task-copy-from-copy-not-supported')
  let names=['Source'],values={Source:1},copies=0,writes=0,renames=0
  const api={
    inspect:async()=>inv(names),
    readRange:async({sheet})=>read(values[sheet]),
    copySheetVerified:async(from,to)=>{copies++;names.push(to);values[to]=values[from];return {ok:true,authority:'LIVE_VERIFY',noOp:false}},
    writeRangeVerified:async({sheet,values:v})=>{writes++;values[sheet]=v[0][0];return {ok:true,authority:'LIVE_VERIFY'}},
    renameSheetVerified:async(from,to)=>{renames++;names=names.map(x=>x===from?to:x);values[to]=values[from];delete values[from];return {ok:true,authority:'LIVE_VERIFY'}},
    createSheetVerified:async()=>{throw new Error('unexpected create')},deleteSheetVerified:async()=>{throw new Error('unexpected delete')}
  }
  const result=await task.executeTask({task:{operations:[{intent:'copy_sheet',sheet:'Source',name:'Copy'},{intent:'write_range',sheet:'Copy',range:'A1',values:[[9]]},{intent:'rename_sheet',sheet:'Copy',name:'Final'}]},api})
  assert.equal(result.ok,true);assert.equal(result.authority,'LIVE_VERIFY');assert.equal(result.noOp,false)
  assert.deepEqual(result.receipt.map(x=>x.status),['APPLIED','APPLIED','APPLIED'])
  assert.equal(copies,1);assert.equal(writes,1);assert.equal(renames,1);assert.deepEqual(names,['Source','Final']);assert.equal(values.Final,9)
  assert.equal(result.wholeTaskVerification.checks.every(x=>x.status==='PASS'),true)
  console.log('XLSX AGENT COPY INTEGRATED STATIC: PASS')
}
main().catch(e=>{console.error(e);process.exitCode=1})
