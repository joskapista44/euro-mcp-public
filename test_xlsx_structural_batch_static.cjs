'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
for(const [intent,range] of [['insert_rows','A2:A3'],['delete_rows','A2'],['insert_columns','B1:C1'],['delete_columns','B1']]){
 const op={intent,sheet:'S',range},p=batch.planTask({operations:[op]});assert.equal(p.ok,true,intent);assert.equal(p.steps[0].family,'structural',intent);assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true,intent)
}
assert.equal(batch.planTask({operations:[{intent:'insert_rows',sheet:'S',range:'B2'}]}).ok,false)
assert.equal(batch.planTask({operations:[{intent:'insert_columns',sheet:'S',range:'B2'}]}).ok,false)
const normalized=batch.planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2'}]}).steps[0].operation;assert.equal(normalized.preconditionFingerprint,null);assert.equal(normalized.expectedPostFingerprint,null)
console.log('XLSX STRUCTURAL BATCH STATIC: PASS')
