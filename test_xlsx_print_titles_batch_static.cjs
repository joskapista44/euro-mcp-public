'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const task={operations:[{intent:'set_print_titles',sheet:'Report',axis:'rows',from:1,to:2},{intent:'set_print_titles',sheet:'Report',axis:'columns',from:1,to:2}]}
const p=batch.planTask(task);assert.equal(p.ok,true);assert.deepEqual(p.steps.map(x=>x.family),['print-titles','print-titles']);assert.equal(mcp.schema.safeParse({file_id:'1',operations:task.operations}).success,true)
assert.equal(batch.planTask({operations:[task.operations[0],task.operations[0]]}).ok,false)
console.log('XLSX PRINT TITLES BATCH STATIC: PASS')
