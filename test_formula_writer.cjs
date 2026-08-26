'use strict'
const assert=require('assert')
const fw=require('./formula-writer.cjs')

let v=fw.validateFormulaMatrix('B2:C3',[['=A1','=A2'],['=$A$1','=SUM(A1:A2)']]); assert.equal(v.ok,true); assert.equal(v.parsed.cellCount,4)
assert.equal(fw.validateFormulaMatrix('A1:A2',[['=1','=2']]).outcome,'dimension-mismatch')
assert.equal(fw.validateFormulaMatrix('A1',[[1]]).outcome,'invalid-formula')
assert.equal(fw.validateFormulaMatrix('A1:Z1001',Array.from({length:1001},()=>Array(26).fill('=1'))).outcome,'range-too-large')

const formulas=[['=A1','=A2'],['=$A$1','=SUM(A1:A2)']]
const cells={}
global.Api={GetSheet:(name)=>name==='Second'?{GetRange:(addr)=>{if(!cells[addr])cells[addr]={SetFormula:(f)=>{cells[addr].formula=f}};return cells[addr]}}:null}
let r=fw.formulaWriterCommand('Second','B2:C3',formulas,26000); assert.equal(r.ok,true); assert.equal(r.writtenFormulas,4); assert.equal(cells.B2.formula,'=A1'); assert.equal(cells.C3.formula,'=SUM(A1:A2)')

const live={ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'=A1'},{dataType:'formula',formula:'=A2'}],[{dataType:'formula',formula:'=$A$1'},{dataType:'formula',formula:'=SUM(A1:A2)'}]]}
let verified=fw.verifyFormulaMatrix(live,formulas); assert.equal(verified.ok,true)
let mismatch=fw.verifyFormulaMatrix({ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'=WRONG'},{dataType:'formula',formula:'=A2'}],[{dataType:'formula',formula:'=$A$1'},{dataType:'formula',formula:'=SUM(A1:A2)'}]]},formulas); assert.equal(mismatch.outcome,'verification-mismatch'); assert.equal(mismatch.mismatches.length,1)
let wrongSource=fw.verifyFormulaMatrix({...live,source:'persisted-webdav'},formulas); assert.equal(wrongSource.outcome,'verification-unavailable')
let unknown=fw.verifyFormulaMatrix({ok:true,source:'live-coedit-editor',cells:[[{dataType:'unknown',formula:null}]]},[['=A1']]); assert.equal(unknown.outcome,'verification-mismatch')

delete global.Api
console.log('test_formula_writer: ok')
