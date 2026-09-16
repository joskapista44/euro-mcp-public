'use strict'
const assert=require('assert')
const c=require('./xlsx-copy-chain-retry-classifier.cjs')
function cell(address,value){return {address,rawValue:value,value,displayText:String(value),formula:null,dataType:typeof value==='number'?'number':'string'}}
function read(sheet,range,values,startRow=1,startCol=1){return {ok:true,authority:'LIVE_READ',sheet,range,rows:values.length,columns:values[0].length,cells:values.map((row,r)=>row.map((v,col)=>cell(`${String.fromCharCode(64+startCol+col)}${startRow+r}`,v)))}}
function inventory(targetValues,targetRange='A1:B2'){return {ok:true,authority:'LIVE_READ',sheets:[{name:'Source',index:0,usedRange:'A1:B2'},{name:'Final',index:1,usedRange:targetRange,_values:targetValues}]}}
const plan={ok:true,operations:[{index:0,intent:'copy_sheet',sheet:'Source',name:'Copy'},{index:1,intent:'write_range',sheet:'Copy',range:'A1',values:[[99]],formulas:null},{index:2,intent:'rename_sheet',sheet:'Copy',name:'Final'}]}
async function classify(values,targetRange='A1:B2',p=plan){const inv=inventory(values,targetRange);const api={readRange:async({sheet,range})=>sheet==='Source'?read(sheet,range,[[11,22],[33,44]]):read(sheet,range,values)};return c.classifyCopyChainRetry({plan:p,api,inventory:inv})}
;(async()=>{
 assert.deepEqual(c.parseRange('$B$2:C3'),{start:{row:2,column:2},end:{row:3,column:3}});assert.equal(c.rangeAddress(c.unionRange([c.parseRange('A1:B2'),c.parseRange('D4')])), 'A1:D4')
 const ok=await classify([[99,22],[33,44]]);assert.equal(ok.matched,true);assert.equal(ok.ok,true);assert.equal(ok.authority,'LIVE_VERIFY');assert.equal(ok.noOp,true);assert.equal(ok.outcome,'xlsx-copy-chain-task-already-satisfied');assert.equal(ok.receipt.length,3);assert.ok(ok.receipt.every(x=>x.status==='NO_OP_SATISFIED'))
 const baseline=await classify([[99,777],[33,44]]);assert.equal(baseline.ok,false);assert.equal(baseline.outcome,'xlsx-copy-chain-retry-copy-baseline-mismatch');assert.notEqual(baseline.authority,'LIVE_VERIFY')
 const write=await classify([[98,22],[33,44]]);assert.equal(write.ok,false);assert.equal(write.outcome,'xlsx-copy-chain-retry-downstream-write-mismatch');assert.notEqual(write.authority,'LIVE_VERIFY')
 const extendedPlan={ok:true,operations:[{index:0,intent:'copy_sheet',sheet:'Source',name:'Copy'},{index:1,intent:'write_range',sheet:'Copy',range:'C3',values:[[77]],formulas:null},{index:2,intent:'rename_sheet',sheet:'Copy',name:'Final'}]};const extended=await classify([[11,22,null],[33,44,null],[null,null,77]],'A1:C3',extendedPlan);assert.equal(extended.ok,false);assert.equal(extended.outcome,'xlsx-copy-chain-retry-unproven-expanded-cell');assert.notEqual(extended.authority,'LIVE_VERIFY')
 const notFinal=await c.classifyCopyChainRetry({plan,api:{},inventory:{ok:true,authority:'LIVE_READ',sheets:[{name:'Source',usedRange:'A1:B2'},{name:'Copy',usedRange:'A1:B2'}]}});assert.equal(notFinal.matched,false)
 console.log('XLSX COPY CHAIN RETRY CLASSIFIER STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
