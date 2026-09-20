'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const ops=[{intent:'set_page_break',sheet:'Report',mode:'add',axis:'row',at:11},{intent:'set_page_break',sheet:'Report',mode:'add',axis:'column',at:5}]
const p=batch.planTask({operations:ops});assert.equal(p.ok,true);assert.deepEqual(p.steps.map(x=>x.family),['page-break','page-break']);assert.equal(mcp.schema.safeParse({file_id:'1',operations:ops}).success,true)
assert.equal(batch.planTask({operations:[ops[0],ops[0]]}).ok,false)
assert.equal(mcp.schema.safeParse({file_id:'1',operations:[{intent:'set_page_break',sheet:'Report',mode:'reset'}]}).success,true)
console.log('XLSX PAGE BREAK BATCH STATIC: PASS')
