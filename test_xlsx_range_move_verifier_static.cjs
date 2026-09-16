'use strict'
const assert=require('assert'),{verifyRangeMoveSemantic}=require('./xlsx-range-move-verifier.cjs')
const read=cells=>({ok:true,authority:'LIVE_READ',cells})
const source=read([[{row:1,column:1,value:5},{row:1,column:2,formula:'=A1+1'}]])
let r=verifyRangeMoveSemantic(source,read([[{value:'',formula:null,dataType:'blank'},{value:'',formula:null,dataType:'blank'}]]),read([[{row:1,column:4,value:5},{row:1,column:5,formula:'=D1+1'}]]));assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY')
r=verifyRangeMoveSemantic(source,read([[{value:5},{value:'',formula:null,dataType:'blank'}]]),read([[{row:1,column:4,value:5},{row:1,column:5,formula:'=D1+1'}]]));assert.equal(r.ok,false);assert.equal(r.outcome,'range-move-source-not-cleared')
r=verifyRangeMoveSemantic(source,read([[{value:'',formula:null,dataType:'blank'},{value:'',formula:null,dataType:'blank'}]]),read([[{row:1,column:4,value:6},{row:1,column:5,formula:'=D1+1'}]]));assert.equal(r.ok,false);assert.equal(r.outcome,'range-move-target-semantic-mismatch')
console.log('XLSX RANGE MOVE VERIFIER STATIC: PASS')
