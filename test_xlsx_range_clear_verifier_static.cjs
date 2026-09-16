'use strict'
const assert=require('assert'),{verifyRangeClearSemantic}=require('./xlsx-range-clear-verifier.cjs')
let r=verifyRangeClearSemantic({ok:true,authority:'LIVE_READ',cells:[[{dataType:'blank',value:null,formula:null},{dataType:'blank',rawValue:null}]]});assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY')
r=verifyRangeClearSemantic({ok:true,authority:'LIVE_READ',cells:[[{dataType:'string',value:'x'}]]});assert.equal(r.ok,false);assert.equal(r.outcome,'range-clear-semantic-mismatch')
r=verifyRangeClearSemantic({ok:true,authority:'LIVE_READ',cells:[[{dataType:'blank',value:null,formula:'=1+1'}]]});assert.equal(r.ok,false)
console.log('XLSX RANGE CLEAR VERIFIER STATIC: PASS')
