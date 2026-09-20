'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
const ops=[
 {intent:'freeze_panes',sheet:'Report',mode:'at',range:'A3'},
 {intent:'set_print_setup',sheet:'Report',mode:'fit_to_pages',fitToWidth:1,fitToHeight:1},
 {intent:'set_print_titles',sheet:'Report',axis:'rows',from:1,to:2},
 {intent:'set_print_titles',sheet:'Report',axis:'columns',from:1,to:1}
]
const parsed=mcp.schema.safeParse({file_id:'1',operations:ops});assert.equal(parsed.success,true)
const p=batch.planTask({operations:ops});assert.equal(p.ok,true);assert.deepEqual(p.steps.map(x=>x.family),['print-setup','print-titles','print-titles','freeze'])
assert.equal(batch.planTask({operations:[ops[2],ops[2]]}).ok,false)
console.log('XLSX PRINT PAGE INTENTS SCHEMA STATIC: PASS')
