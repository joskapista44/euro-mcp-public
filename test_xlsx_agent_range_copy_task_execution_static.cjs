'use strict'
const assert=require('assert'),task=require('./xlsx-agent-range-copy-task.cjs')
const inventory={ok:true,authority:'LIVE_READ',sheets:[{name:'S'}]}
const cells=[[{row:1,column:1,value:7}]]
let inspections=0,copies=0
const api={inspect:async()=>{inspections++;return inventory},readRange:async()=>({ok:true,authority:'LIVE_READ',cells}),copyRangeVerified:async()=>{copies++;return {ok:true,authority:'LIVE_VERIFY',noOp:false}}}
;(async()=>{let r=await task.executeRangeCopyTask({task:{operations:[{intent:'copy_range',sheet:'S',range:'A1',targetRange:'D1'}]},api});assert.equal(r.ok,true);assert.equal(r.noOp,true);assert.equal(copies,0);assert.equal(r.receipt[0].status,'NO_OP_SATISFIED');assert.equal(r.wholeTaskVerification.ok,true);assert.ok(inspections>=2);console.log('XLSX AGENT RANGE COPY TASK EXECUTION STATIC: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
