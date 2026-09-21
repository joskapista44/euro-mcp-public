'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs'),mcp=require('./xlsx-batch-mcp.cjs')
for(const [intent,range] of [['insert_rows','A2:A3'],['delete_rows','A2:A2'],['insert_columns','B1:C1'],['delete_columns','B1:B1']]){
 const op={intent,sheet:'S',range},p=batch.planTask({operations:[op]});assert.equal(p.ok,true,intent);assert.equal(p.steps[0].family,'structural',intent);assert.equal(mcp.schema.safeParse({file_id:'123',operations:[op]}).success,true,intent)
}
assert.equal(batch.planTask({operations:[{intent:'insert_rows',sheet:'S',range:'B2'}]}).ok,false)
assert.equal(batch.planTask({operations:[{intent:'insert_columns',sheet:'S',range:'B2'}]}).ok,false)
const normalized=batch.planTask({operations:[{intent:'insert_rows',sheet:'S',range:'A2:A2'}]}).steps[0].operation;assert.equal(normalized.preconditionFingerprint,null);assert.equal(normalized.expectedPostFingerprint,null)
const rowProjection=batch.planTask({operations:[
 {intent:'create_sheet',name:'S'},
 {intent:'write_range',sheet:'S',range:'A1:C4',values:[['h1','h2','h3'],['a','b','c'],['d','e','f'],['g','h','i']],formulas:[[null,null,null],[null,null,'=A2'],[null,null,'=A3'],[null,null,'=A4']]},
 {intent:'insert_rows',sheet:'S',range:'A2:C2'}
]})
assert.equal(rowProjection.ok,true)
assert.equal(rowProjection.steps[0].operations[1].range,'A1:C4')
assert.deepEqual(rowProjection.steps[0].verifyOperations.slice(1).map(op=>op.range),['A1:C1','A3:C5'])
assert.deepEqual(rowProjection.steps[0].verifyOperations[2].formulas.map(row=>row[2]),['=A3','=A4','=A5'])
const columnProjection=batch.planTask({operations:[
 {intent:'create_sheet',name:'C'},
 {intent:'write_range',sheet:'C',range:'A1:C2',values:[[1,2,3],[4,5,6]],formulas:[[null,'=A1','=B1'],[null,'=A2','=B2']]},
 {intent:'insert_columns',sheet:'C',range:'B1:B2'}
]})
assert.equal(columnProjection.ok,true)
assert.deepEqual(columnProjection.steps[0].verifyOperations.slice(1).map(op=>op.range),['A1:A2','C1:D2'])
assert.deepEqual(columnProjection.steps[0].verifyOperations[2].formulas,[['=A1','=C1'],['=A2','=C2']])
console.log('XLSX STRUCTURAL BATCH STATIC: PASS')
