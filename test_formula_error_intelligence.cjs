'use strict'
const assert=require('assert')
const fs=require('fs')
const fe=require('./formula-error-intelligence.cjs')

function cell(address,formulaStatus,formula,value,text,cellType='formula'){
  return {address,row:Number(address.match(/\d+/)[0]),column:1,formulaStatus,formula,calculatedValue:value,displayText:text,cellType}
}
function inspected(cells,unsupported=[]){
  return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:'Calc',range:'A1:A'+cells.length,rows:cells.length,columns:1,cellCount:cells.length,cells:cells.map(x=>[x]),summary:{},unsupported}
}

assert.strictEqual(fe.normalizeErrorToken('#ref!'),'#REF!')
assert.strictEqual(fe.normalizeErrorToken(' #DIV/0! '),'#DIV/0!')
assert.strictEqual(fe.normalizeErrorToken('#NOTREAL!'),null)

let result=fe.diagnoseInspection(inspected([
  cell('A1','formula','=1/0','#DIV/0!','#DIV/0!'),
  cell('A2','formula','=Missing!A1','#REF!','#REF!'),
  cell('A3','formula','=UNKNOWNFUNC(1)','#NAME?','#NAME?'),
  cell('A4','formula','=VALUE("x")','#VALUE!','#VALUE!'),
  cell('A5','formula','=1+1',2,'2'),
  cell('A6','constant',null,'#REF!','#REF!','string'),
  cell('A7','constant',null,'','','blank')
]))
assert.strictEqual(result.ok,true)
assert.strictEqual(result.summary.formulaErrors,4)
assert.strictEqual(result.summary.formulaOk,1)
assert.strictEqual(result.summary.notFormula,2)
assert.deepStrictEqual(result.summary.byType,{ 'DIV/0':1, REF:1, NAME:1, VALUE:1 })
assert.deepStrictEqual(result.errors.map(x=>[x.address,x.errorType]),[['A1','DIV/0'],['A2','REF'],['A3','NAME'],['A4','VALUE']])
assert.strictEqual(result.cells[5][0].errorStatus,'not-formula')

result=fe.diagnoseInspection(inspected([
  cell('A1','formula','=NA()','#N/A','#N/A'),
  cell('A2','formula','=SQRT(-1)','#NUM!','#NUM!'),
  cell('A3','formula','=A1:A2','#SPILL!','#SPILL!')
]))
assert.deepStrictEqual(result.errors.map(x=>x.errorType),['N/A','NUM','SPILL'])

result=fe.diagnoseInspection(inspected([
  cell('A1','formula','=1/0',null,null)
],[{field:'GetValue',reason:'unavailable'},{field:'GetText',reason:'unavailable'}]))
assert.strictEqual(result.cells[0][0].errorStatus,'unknown')
assert.strictEqual(result.summary.unknown,1)

result=fe.diagnoseInspection({ok:false,outcome:'sheet-not-found',source:'live-coedit-editor'})
assert.strictEqual(result.outcome,'sheet-not-found')
result=fe.diagnoseInspection({ok:true,cells:null})
assert.strictEqual(result.outcome,'unsupported-result')

const source=fs.readFileSync(require.resolve('./formula-error-intelligence.cjs'),'utf8')+fs.readFileSync(require.resolve('./euro-mcp-m35.cjs'),'utf8')
assert.ok(source.includes('formulaInspector.inspectFormulaInFrame'))
assert.ok(source.includes('spreadsheeteditor'))
assert.ok(source.includes('callCommand'))
for(const forbidden of ['DocBuilder.OpenFile','JSZip','adm-zip','unzipper','xlsx.readFile','readFileSync(fileId']) assert.ok(!source.includes(forbidden),`forbidden offline fallback marker: ${forbidden}`)

console.log('M3.5 formula error intelligence tests: PASS')
