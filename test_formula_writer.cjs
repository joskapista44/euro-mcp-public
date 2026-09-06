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
let r=fw.formulaWriterCommand('Second','B2:C3',formulas,26000); assert.equal(r.ok,true); assert.equal(r.writtenFormulas,4); assert.equal(r.writer,'SetFormula'); assert.equal(cells.B2.formula,'=A1'); assert.equal(cells.C3.formula,'=SUM(A1:A2)')

const valueCells={}
global.Api={GetSheet:(name)=>name==='Second'?{GetRange:(addr)=>{if(!valueCells[addr])valueCells[addr]={SetValue:(f)=>{valueCells[addr].formula=f}};return valueCells[addr]}}:null}
r=fw.formulaWriterCommand('Second','B2:C3',formulas,26000); assert.equal(r.ok,true); assert.equal(r.writtenFormulas,4); assert.equal(r.writer,'SetValue'); assert.equal(valueCells.B2.formula,'=A1'); assert.equal(valueCells.C3.formula,'=SUM(A1:A2)')

global.Api={GetSheet:(name)=>name==='Second'?{GetRange:()=>({})}:null}
r=fw.formulaWriterCommand('Second','A1',[['=1']],26000); assert.equal(r.ok,false); assert.equal(r.outcome,'unsupported')

// Measured live ONLYOFFICE canonicalization: optional quotes around a simple sheet name may be
// removed and whitespace may be inserted outside string literals.
assert.equal(fw.canonicalizeFormula("='M44S_123'!A1*2"),'=M44S_123!A1*2')
assert.equal(fw.canonicalizeFormula('= M44S_123!A1*2'),'=M44S_123!A1*2')
// Do not erase semantically significant whitespace inside Excel string literals.
assert.notEqual(fw.canonicalizeFormula('="a b"'),fw.canonicalizeFormula('="ab"'))
assert.equal(fw.canonicalizeFormula('="a b" & A1'),'="a b"&A1')
// Quoting remains significant for sheet names that cannot be bare identifiers.
assert.equal(fw.canonicalizeFormula("='Sales 2026'!A1"),"='Sales 2026'!A1")
assert.equal(fw.canonicalizeFormula("='O''Brien'!A1"),"='O''Brien'!A1")

const live={ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'=A1'},{dataType:'formula',formula:'=A2'}],[{dataType:'formula',formula:'=$A$1'},{dataType:'formula',formula:'=SUM(A1:A2)'}]]}
let verified=fw.verifyFormulaMatrix(live,formulas); assert.equal(verified.ok,true)

const canonicalLive={ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'= M44S_123!A1*2'}]]}
verified=fw.verifyFormulaMatrix(canonicalLive,[["='M44S_123'!A1*2"]]); assert.equal(verified.ok,true)

let mismatch=fw.verifyFormulaMatrix({ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'=WRONG'},{dataType:'formula',formula:'=A2'}],[{dataType:'formula',formula:'=$A$1'},{dataType:'formula',formula:'=SUM(A1:A2)'}]]},formulas); assert.equal(mismatch.outcome,'verification-mismatch'); assert.equal(mismatch.mismatches.length,1)
let stringMismatch=fw.verifyFormulaMatrix({ok:true,source:'live-coedit-editor',cells:[[{dataType:'formula',formula:'="ab"'}]]},[['="a b"']]); assert.equal(stringMismatch.outcome,'verification-mismatch')
let wrongSource=fw.verifyFormulaMatrix({...live,source:'persisted-webdav'},formulas); assert.equal(wrongSource.outcome,'verification-unavailable')
let unknown=fw.verifyFormulaMatrix({ok:true,source:'live-coedit-editor',cells:[[{dataType:'unknown',formula:null}]]},[['=A1']]); assert.equal(unknown.outcome,'verification-mismatch')

delete global.Api
console.log('test_formula_writer: ok')
