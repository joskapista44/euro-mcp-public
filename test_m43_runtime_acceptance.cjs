'use strict'
const assert = require('assert')
const fs = require('fs')
const { statusOf, overall, markerVerification, originalContentVerification } = require('./m43-runtime-acceptance.cjs')

assert.strictEqual(statusOf({ok:true,verification:{status:'PASS'}}),'PASS')
assert.strictEqual(statusOf({ok:false,verification:{status:'UNKNOWN'}}),'UNKNOWN')
assert.strictEqual(statusOf({ok:false,verification:{status:'FAIL'}}),'FAIL')
assert.strictEqual(overall([{status:'PASS'},{status:'PASS'}]),'PASS')
assert.strictEqual(overall([{status:'PASS'},{status:'UNKNOWN'}]),'UNKNOWN')
assert.strictEqual(overall([{status:'UNKNOWN'},{status:'FAIL'}]),'FAIL')

assert.strictEqual(markerVerification({ok:true,cells:[[{address:'J21',rawValue:'MARK'}]]},'J21','MARK').status,'PASS')
assert.strictEqual(markerVerification({ok:true,cells:[[{address:'J21',rawValue:'OTHER'}]]},'J21','MARK').status,'FAIL')
assert.strictEqual(markerVerification({ok:false},'J21','MARK').status,'FAIL')

assert.strictEqual(
  originalContentVerification(
    {ok:true,cells:[[{address:'J20',rawValue:42,value:42,displayText:'42',formula:null}]]},
    'J20',
    {address:'J20',rawValue:42,formula:null},
  ).status,
  'PASS',
)
assert.strictEqual(
  originalContentVerification(
    {ok:true,cells:[[{address:'J20',rawValue:99,value:99,displayText:'99',formula:null}]]},
    'J20',
    {address:'J20',rawValue:42,formula:null},
  ).status,
  'FAIL',
)
assert.strictEqual(
  originalContentVerification(
    {ok:true,cells:[[{address:'J20',rawValue:2,value:2,displayText:'2',formula:'=1+1'}]]},
    'J20',
    {address:'J20',rawValue:2,formula:'=1+1'},
  ).status,
  'PASS',
)
assert.strictEqual(
  originalContentVerification(
    {ok:true,cells:[[{address:'J20',rawValue:3,value:3,displayText:'3',formula:'=1+2'}]]},
    'J20',
    {address:'J20',rawValue:2,formula:'=1+1'},
  ).status,
  'FAIL',
)

const src=fs.readFileSync('m43-runtime-acceptance.cjs','utf8')
assert.match(src,/runLayoutInFrame/)
assert.match(src,/runOperationInFrame/)
assert.match(src,/readRangeInFrame/)
assert.match(src,/writeBulkInFrame/)
assert.match(src,/column-width/)
assert.match(src,/row-height/)
assert.match(src,/column-hide/)
assert.match(src,/column-show/)
assert.match(src,/row-hide/)
assert.match(src,/row-show/)
assert.match(src,/autofit-column/)
assert.match(src,/autofit-row/)
assert.match(src,/structuralRow.*20:20/)
assert.match(src,/structuralColumn.*J:J/)
assert.match(src,/rowMovedCell.*J21/)
assert.match(src,/columnMovedCell.*K20/)
assert.match(src,/marker-seed-readback/)
assert.match(src,/structural-mutation-gate/)
assert.match(src,/column-structural-gate/)
assert.match(src,/marker-restore-gate/)
assert.match(src,/marker-restore-readback/)
assert.match(src,/originalContentVerification/)
assert.match(src,/row-insert/)
assert.match(src,/row-delete/)
assert.match(src,/column-insert/)
assert.match(src,/column-delete/)
assert.match(src,/humanObservationRequired: false/)
assert.doesNotMatch(src,/PROPFIND|unzip|writeFile/i)
console.log('test_m43_runtime_acceptance: OK')
