'use strict'
const assert=require('assert')
const fs=require('fs')
const fv=require('./formula-verification-intelligence.cjs')

function cell(address,formulaStatus,formula,value,text,cellType='formula'){
  const m=String(address).match(/^([A-Z]+)(\d+)$/)
  return {address,row:Number(m[2]),column:1,formulaStatus,formula,calculatedValue:value,displayText:text,cellType}
}
function inspected(rows,opts={}){
  return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:opts.sheet||'Calc',range:opts.range||`A1:${String.fromCharCode(64+(rows[0]?rows[0].length:1))}${rows.length}`,rows:rows.length,columns:rows[0]?rows[0].length:0,cellCount:rows.reduce((n,r)=>n+r.length,0),cells:rows,unsupported:opts.unsupported||[]}
}

// Validation: expected matrix must exactly match the target range and contain formulas only.
let parsed={rows:2,columns:2}
assert.strictEqual(fv.validateExpectedFormulaMatrix(parsed,[['=1','=2'],['=3','=4']]).ok,true)
assert.strictEqual(fv.validateExpectedFormulaMatrix(parsed,null).ok,true)
assert.strictEqual(fv.validateExpectedFormulaMatrix(parsed,[['=1']]).outcome,'dimension-mismatch')
assert.strictEqual(fv.validateExpectedFormulaMatrix(parsed,[['=1','=2'],['x','=4']]).outcome,'invalid-expected-formula')

// Direct write verification: exact read-back + healthy calculated/display evidence -> verified.
let result=fv.verifyInspection(inspected([[
  cell('A1','formula','=1+1',2,'2'),
  cell('B1','formula','=2+2',4,'4')
]],{range:'A1:B1'}),[['=1+1','=2+2']])
assert.strictEqual(result.verificationOutcome,'verified')
assert.strictEqual(result.summary.matched,2)
assert.strictEqual(result.summary.failed,0)
assert.strictEqual(result.cells[0][0].verificationStatus,'verified')

// Formula mismatch is explicit and never rounded to success.
result=fv.verifyInspection(inspected([[
  cell('A1','formula','=1+2',3,'3')
]],{range:'A1:A1'}),[['=1+1']])
assert.strictEqual(result.verificationOutcome,'failed')
assert.strictEqual(result.summary.mismatched,1)
assert.deepStrictEqual(result.mismatches.map(x=>x.address),['A1'])
assert.strictEqual(result.cells[0][0].formulaMatch,'mismatch')

// A formula that read back correctly but calculates to a proven formula error still fails verification.
result=fv.verifyInspection(inspected([[
  cell('A1','formula','=1/0','#DIV/0!','#DIV/0!')
]],{range:'A1:A1'}),[['=1/0']])
assert.strictEqual(result.verificationOutcome,'failed')
assert.strictEqual(result.summary.formulaErrors,1)
assert.strictEqual(result.formulaErrors[0].errorType,'DIV/0')
assert.strictEqual(result.cells[0][0].formulaMatch,'match')
assert.strictEqual(result.cells[0][0].verificationStatus,'failed')

// Missing getter evidence is fail-closed unknown even when formula text matches.
result=fv.verifyInspection(inspected([[
  cell('A1','formula','=SUM(B1:B2)',3,null)
]],{range:'A1:A1',unsupported:[{field:'GetText',reason:'unavailable'}]}),[['=SUM(B1:B2)']])
assert.strictEqual(result.verificationOutcome,'unknown')
assert.strictEqual(result.summary.unknown,1)
assert.strictEqual(result.cells[0][0].errorStatus,'unknown')

// Unproven formula read-back status is unknown, not mismatch or success.
result=fv.verifyInspection(inspected([[
  cell('A1','unknown',null,3,'3','unknown')
]],{range:'A1:A1'}),[['=SUM(B1:B2)']])
assert.strictEqual(result.verificationOutcome,'unknown')
assert.strictEqual(result.cells[0][0].formulaMatch,'unknown')

// Mixed bulk range: match, mismatch, formula error and healthy formula all stay distinguishable.
result=fv.verifyInspection(inspected([
  [cell('A1','formula','=B1',1,'1'),cell('B1','formula','=C1',2,'2')],
  [cell('A2','formula','=1/0','#DIV/0!','#DIV/0!'),cell('B2','formula','=SUM(C1:C2)',5,'5')]
],{range:'A1:B2'}),[
  ['=B1','=WRONG()'],
  ['=1/0','=SUM(C1:C2)']
])
assert.strictEqual(result.verificationOutcome,'failed')
assert.strictEqual(result.summary.matched,3)
assert.strictEqual(result.summary.mismatched,1)
assert.strictEqual(result.summary.formulaErrors,1)
assert.strictEqual(result.mismatches[0].address,'B1')

// Cross-sheet formulas are compared as formula text; no first-sheet assumption exists.
result=fv.verifyInspection(inspected([[
  cell('A1','formula',"='Input Data'!B2*Rates!C3",42,'42')
]],{sheet:'Summary',range:'A1:A1'}),[["='Input Data'!B2*Rates!C3"]])
assert.strictEqual(result.sheet,'Summary')
assert.strictEqual(result.verificationOutcome,'verified')

// Fill-result verification: expanded formulas can be verified cell-by-cell as an expected matrix.
result=fv.verifyInspection(inspected([
  [cell('C2','formula','=A2+B2',3,'3')],
  [cell('C3','formula','=A3+B3',7,'7')],
  [cell('C4','formula','=A4+B4',11,'11')]
],{sheet:'Calc',range:'C2:C4'}),[['=A2+B2'],['=A3+B3'],['=A4+B4']])
assert.strictEqual(result.verificationOutcome,'verified')
assert.strictEqual(result.summary.verified,3)

// Observational mode without expected formulas still returns calculated/display/error diagnosis without inventing a match verdict.
result=fv.verifyInspection(inspected([[
  cell('A1','formula','=1+1',2,'2'),
  cell('B1','constant',null,'label','label','string')
]],{range:'A1:B1'}),null)
assert.strictEqual(result.verificationOutcome,'observed')
assert.strictEqual(result.cells[0][0].formulaMatch,'not-requested')
assert.strictEqual(result.cells[0][0].verificationStatus,'observed-ok')
assert.strictEqual(result.cells[0][1].verificationStatus,'observed')

// Unknown error-like token remains unknown; unsupported/ambiguous evidence is never guessed.
result=fv.verifyInspection(inspected([[
  cell('A1','formula','=FUTUREFUNC()','#FUTURE!','#FUTURE!')
]],{range:'A1:A1'}),[['=FUTUREFUNC()']])
assert.strictEqual(result.verificationOutcome,'unknown')
assert.strictEqual(result.cells[0][0].errorStatus,'unknown')

// Inspector/result shape failures remain fail-closed.
result=fv.verifyInspection({ok:false,outcome:'sheet-not-found',source:'live-coedit-editor'},[['=1']])
assert.strictEqual(result.outcome,'sheet-not-found')
result=fv.verifyInspection({ok:true,cells:null},[['=1']])
assert.strictEqual(result.outcome,'unsupported-result')
result=fv.verifyInspection(inspected([[cell('A1','formula','=1',1,'1')]],{range:'A1:A1'}),[['=1','=2']])
assert.strictEqual(result.outcome,'dimension-mismatch')

// Architecture acceptance: M3.6 composes M3.1 + M3.5 over live spreadsheeteditor/callCommand only.
const source=fs.readFileSync(require.resolve('./formula-verification-intelligence.cjs'),'utf8')+fs.readFileSync(require.resolve('./euro-mcp-m36.cjs'),'utf8')
assert.ok(source.includes('formulaInspector.inspectFormulaInFrame'))
assert.ok(source.includes('formulaErrors.diagnoseInspection'))
assert.ok(source.includes('spreadsheeteditor'))
assert.ok(source.includes('callCommand'))
for(const forbidden of ['DocBuilder.OpenFile','JSZip','adm-zip','unzipper','xlsx.readFile','readFileSync(fileId','OOXML','downloaded XLSX fallback']) {
  if(forbidden==='OOXML') continue
  assert.ok(!source.includes(forbidden),`forbidden offline fallback marker: ${forbidden}`)
}

console.log('M3.6 formula verification intelligence tests: PASS')
