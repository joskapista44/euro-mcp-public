'use strict'
const assert=require('assert'),task=require('./xlsx-agent-range-copy-task.cjs')
let r=task.planTask({operations:[{intent:'copy_range',sheet:'S',range:'A1:B2',targetRange:'D1:E2'}]});assert.equal(r.ok,true);assert.equal(r.operation.targetSheet,'S')
r=task.planTask({operations:[{intent:'copy_range',sheet:'S',range:'A1:B2',targetSheet:'T',targetRange:'D1:E2'}]});assert.equal(r.ok,true);assert.equal(r.operation.targetSheet,'T')
r=task.planTask({operations:[{intent:'copy_range',sheet:'S',range:'',targetRange:'D1'}]});assert.equal(r.ok,false);assert.equal(r.authority,'PLAN_ONLY')
console.log('XLSX AGENT RANGE COPY TASK STATIC: PASS')
