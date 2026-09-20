'use strict'
const assert=require('assert/strict'),mcp=require('./xlsx-batch-mcp.cjs'),batch=require('./xlsx-persistent-batch.cjs')
const op={intent:'set_page_layout',sheet:'Report',orientation:'xlLandscape',topMargin:10,bottomMargin:10,leftMargin:8,rightMargin:8,printGridlines:false,printHeadings:false}
let p=batch.planTask({operations:[op]});assert.equal(p.ok,true);assert.equal(p.steps.length,1);assert.equal(p.steps[0].family,'page-layout')
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true)
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[{...op,fitToWidth:1}]}).success,false)
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op,{...op}]}).success,true)
p=batch.planTask({operations:[op,{...op}]});assert.equal(p.ok,false);assert.equal(p.outcome,'xlsx-batch-conflicting-goals')
console.log('XLSX PAGE LAYOUT BATCH STATIC: PASS')
