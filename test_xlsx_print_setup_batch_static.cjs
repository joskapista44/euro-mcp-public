'use strict'
const assert=require('assert/strict'),mcp=require('./xlsx-batch-mcp.cjs'),batch=require('./xlsx-persistent-batch.cjs')
const fit={intent:'set_print_setup',sheet:'Report',mode:'fit_to_pages',fitToWidth:1,fitToHeight:1}
const scale={intent:'set_print_setup',sheet:'Report',mode:'scale',scale:85}
const size={intent:'set_print_setup',sheet:'Report',mode:'page_size',width:210,height:297}
for(const op of [fit,scale,size]){assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true);const p=batch.planTask({operations:[op]});assert.equal(p.ok,true);assert.equal(p.steps[0].family,'print-setup')}
assert.equal(mcp.schema.safeParse({file_id:'123',operations:[{...fit,fitToWidth:-1}]}).success,false)
assert.equal(batch.planTask({operations:[fit,{...fit}]}).ok,false)
const multi=batch.planTask({operations:[fit,scale,size]});assert.equal(multi.ok,true);assert.equal(multi.steps.length,3)
console.log('XLSX PRINT SETUP BATCH STATIC: PASS')
