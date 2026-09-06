'use strict'

const assert = require('assert')
const { overall, formulaIntegrity } = require('./m44-runtime-acceptance.cjs')

assert.equal(overall([{status:'PASS'},{status:'PASS'}]), 'PASS')
assert.equal(overall([{status:'PASS'},{status:'UNKNOWN'}]), 'UNKNOWN')
assert.equal(overall([{status:'PASS'},{status:'FAIL'}]), 'FAIL')

const good = formulaIntegrity({
  ok: true,
  cells: [[{ address:'A1', formula:"='M44R_1'!A1*2", rawValue:74, displayText:'74' }]],
}, 'M44R_1', 74)
assert.equal(good.status, 'PASS')

const stale = formulaIntegrity({
  ok: true,
  cells: [[{ address:'A1', formula:"='M44S_1'!A1*2", rawValue:74, displayText:'74' }]],
}, 'M44R_1', 74)
assert.equal(stale.status, 'FAIL')

const ref = formulaIntegrity({
  ok: true,
  cells: [[{ address:'A1', formula:'=#REF!*2', rawValue:'#REF!', displayText:'#REF!' }]],
}, 'M44R_1', 74)
assert.equal(ref.status, 'FAIL')

console.log('test_m44_runtime_acceptance.cjs: OK')
