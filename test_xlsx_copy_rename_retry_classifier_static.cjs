'use strict'
const assert=require('assert')
const c=require('./xlsx-copy-chain-retry-classifier.cjs')
function cell(address,value,formula=null){const blank=(value===null||value===undefined||value==='')&&!formula;return {address,rawValue:blank?null:value,value:blank?null:value,displayText:blank?'':String(value??''),formula,dataType:blank?'blank':typeof value==='number'?'number':'string'}}
function read(sheet,range,values){return {ok:true,authority:'LIVE_READ',sheet,range,rows:values.length,columns:values[0].length,cells:values.map((row,r)=>row.map((v,col)=>cell(`${String.fromCharCode(65+col)}${r+1}`,v)))}}
const plan={ok:true,operations:[{index:0,intent:'copy_sheet',sheet:'Source',name:'Copy'},{index:1,intent:'rename_sheet',sheet:'Copy',name:'Final'}]}
async function classify(targetValues,targetUsed='A1:B2'){
 const inventory={ok:true,authority:'LIVE_READ',sheets:[{name:'Source',index:0,usedRange:'A1:B2'},{name:'Final',index:1,usedRange:targetUsed}]}
 const api={readRange:async({sheet,range})=>sheet==='Source'?read(sheet,range,[[11,22],[33,44]]):read(sheet,range,targetValues)}
 return c.classifyCopyChainRetry({plan,api,inventory})
}
;(async()=>{
 const ok=await classify([[11,22],[33,44]])
 assert.equal(ok.matched,true);assert.equal(ok.ok,true);assert.equal(ok.authority,'LIVE_VERIFY');assert.equal(ok.noOp,true);assert.equal(ok.outcome,'xlsx-copy-chain-task-already-satisfied');assert.deepEqual(ok.retryClassification.downstreamWrites,[]);assert.equal(ok.retryClassification.sourceRange,'A1:B2');assert.equal(ok.retryClassification.targetRange,'A1:B2');assert.equal(ok.receipt.length,2);assert.ok(ok.receipt.every(x=>x.status==='NO_OP_SATISFIED'))
 const corrupt=await classify([[11,22],[33,45]])
 assert.equal(corrupt.matched,true);assert.equal(corrupt.ok,false);assert.equal(corrupt.outcome,'xlsx-copy-chain-retry-copy-baseline-mismatch');assert.notEqual(corrupt.authority,'LIVE_VERIFY')
 const wrongExtent=await classify([[11,22,null],[33,44,null]],'A1:C2')
 assert.equal(wrongExtent.matched,true);assert.equal(wrongExtent.ok,false);assert.equal(wrongExtent.outcome,'xlsx-copy-chain-retry-used-range-mismatch');assert.notEqual(wrongExtent.authority,'LIVE_VERIFY')
 console.log('XLSX COPY RENAME RETRY CLASSIFIER STATIC: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
