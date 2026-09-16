'use strict'
const assert=require('assert'),{verifyRangeCopySemantic}=require('./xlsx-range-copy-verifier.cjs')
const read=cells=>({ok:true,authority:'LIVE_READ',cells})
let r=verifyRangeCopySemantic(read([[{value:1},{formula:'=A1+1'}]]),read([[{value:1},{formula:'A1+1'}]]));assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY')
r=verifyRangeCopySemantic(read([[{value:1}]]),read([[{value:2}]]));assert.equal(r.ok,false);assert.equal(r.outcome,'range-copy-semantic-mismatch')
r=verifyRangeCopySemantic(read([[{value:1},{value:2}]]),read([[{value:1}]]));assert.equal(r.ok,false);assert.equal(r.outcome,'range-copy-shape-mismatch')
console.log('XLSX RANGE COPY VERIFIER STATIC: PASS')
