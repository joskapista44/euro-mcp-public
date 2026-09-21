'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const op={intent:'copy_range',sheet:'Source',range:'A1:B2',targetSheet:'Target',targetRange:'D4:E5'}
const p=batch.planTask({operations:[op]});assert.equal(p.ok,true);assert.equal(p.steps.length,1);assert.equal(p.steps[0].family,'range-copy')
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true)
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[{...op,retryToken:{}}]}).success,false)
console.log('XLSX RANGE COPY BATCH STATIC: PASS')
