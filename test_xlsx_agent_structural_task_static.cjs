'use strict'
const assert=require('assert'),{planTask}=require('./xlsx-agent-structural-task.cjs')
let r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2:A3'}]});assert.equal(r.ok,true);assert.equal(r.operation.type,'rows.insert');assert.equal(r.operation.count,2);assert.equal(r.operation.anchor,2);assert.equal(r.operation.preconditionFingerprint,null)
r=planTask({operations:[{intent:'delete_columns',sheet:'S',range:'C1:E1',preconditionFingerprint:'A'.repeat(64)}]});assert.equal(r.ok,true);assert.equal(r.operation.type,'columns.delete');assert.equal(r.operation.count,3);assert.equal(r.operation.anchor,3);assert.equal(r.operation.preconditionFingerprint,'a'.repeat(64))
r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'B2:B3'}]});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-structural-rows-require-column-a-anchor')
r=planTask({operations:[{intent:'insert_columns',sheet:'S',range:'C2:E2'}]});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-structural-columns-require-row-1-anchor')
r=planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2',preconditionFingerprint:'bad'}]});assert.equal(r.ok,false);assert.equal(r.outcome,'xlsx-structural-task-invalid-precondition-fingerprint')
r=planTask({operations:[{intent:'unknown',sheet:'S',range:'A1'}]});assert.equal(r.ok,false);assert.equal(r.authority,'PLAN_ONLY')
console.log('XLSX AGENT STRUCTURAL TASK STATIC: PASS')
