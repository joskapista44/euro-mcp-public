'use strict'
const assert=require('assert/strict')
const batch=require('./xlsx-persistent-batch.cjs')
const ops=[
 {index:0,intent:'create_sheet',name:'S'},
 {index:1,intent:'write_range',sheet:'S',range:'A1:B4',values:[['id','key'],['r1','C'],['r2','A'],['r3','B']],formulas:[[null,null],[null,null],[null,null],[null,null]]}
]
const tail=[{intent:'sort_range',sheet:'S',range:'A1:B4',keyRange:'B1:B4',order:'asc',hasHeaders:true}]
const p=batch.coreFinalOperations(ops,tail)
assert.deepEqual(p[1].values,[['id','key'],['r2','A'],['r3','B'],['r1','C']])
assert.deepEqual(p[1].formulas,[[null,null],[null,null],[null,null],[null,null]])
assert.deepEqual(ops[1].values,[['id','key'],['r1','C'],['r2','A'],['r3','B']])
const desc=batch.coreFinalOperations(ops,[{...tail[0],order:'desc'}])
assert.deepEqual(desc[1].values,[['id','key'],['r1','C'],['r3','B'],['r2','A']])
const formulaOps=[
 {index:0,intent:'create_sheet',name:'F'},
 {index:1,intent:'write_range',sheet:'F',range:'A1:C4',values:[['id','key','delta'],['r1','C',null],['r2','A',null],['r3','B',null]],formulas:[[null,null,null],[null,null,'=A2+$B$2'],[null,null,'=A3+$B$2'],[null,null,'=A4+$B$2']]}
]
const formulaProjection=batch.coreFinalOperations(formulaOps,[{intent:'sort_range',sheet:'F',range:'A1:C4',keyRange:'B1:B4',order:'asc',hasHeaders:true}])
assert.deepEqual(formulaProjection[1].formulas.map(row=>row[2]),[null,'=A2+$B$2','=A3+$B$2','=A4+$B$2'])
console.log('XLSX BATCH CORE SORT PROJECTION STATIC: PASS')
