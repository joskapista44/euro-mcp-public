'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
for(const type of ['columns.autofit','rows.autofit']){
 const op={intent:'layout_range',sheet:'S',range:'A1:C3',type}
 const p=batch.planTask({operations:[op]});assert.equal(p.ok,true,type);assert.equal(p.steps[0].family,'layout',type)
 assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true,type)
}
console.log('XLSX AUTOFIT BATCH STATIC: PASS')
