'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const ops=[{intent:'set_print_area',sheet:'Report',mode:'set',range:'A1:H20'}]
const p=batch.planTask({operations:ops});assert.equal(p.ok,true);assert.equal(p.steps[0].family,'print-area')
assert.equal(mcp.schema.safeParse({file_id:'1',operations:[{intent:'set_print_area',sheet:'Report',printAreaMode:'set',range:'A1:H20'}]}).success,true)
assert.equal(mcp.schema.safeParse({file_id:'1',operations:[{intent:'set_print_area',sheet:'Report',printAreaMode:'clear'}]}).success,true)
assert.equal(batch.planTask({operations:[ops[0],{...ops[0],mode:'clear',range:undefined}]}).ok,false)
console.log('XLSX PRINT AREA BATCH STATIC: PASS')
