'use strict'

const assert = require('assert')
const fs = require('fs')
const inspector = require('./formula-inspector.cjs')

let passed = 0
function test(name, fn) {
  try { fn(); passed += 1; console.log(`ok - ${name}`) }
  catch (err) { console.error(`not ok - ${name}`); throw err }
}
function withApi(api, fn) {
  const old = global.Api; global.Api = api
  try { return fn() } finally { if (old === undefined) delete global.Api; else global.Api = old }
}

function sheetFor(rangeAddress, formulas, values, texts) {
  return { GetRange(address) {
    if (address !== rangeAddress) throw new Error(`unexpected range ${address}`)
    return { GetFormula:()=>formulas, GetValue:()=>values, GetText:()=>texts }
  } }
}

test('mixed range distinguishes formulas, constants and blanks and returns calculated results', () => {
  let requested = null
  const sheet = sheetFor('A1:C2', [['10','=A1*2','hello'],['','=B1+5','=SUM(A1:B1)']], [[10,20,'hello'],['',25,30]], [['10','20','hello'],['','25','30']])
  const result = withApi({GetSheet(name){requested=name; return name === 'Second Sheet' ? sheet : null}}, () => inspector.formulaInspectorCommand('Second Sheet','A1:C2',26000))
  assert.strictEqual(result.ok,true)
  assert.strictEqual(requested,'Second Sheet')
  assert.strictEqual(result.cells[0][0].formulaStatus,'constant')
  assert.strictEqual(result.cells[0][1].formulaStatus,'formula')
  assert.strictEqual(result.cells[0][1].formula,'=A1*2')
  assert.strictEqual(result.cells[0][1].calculatedValue,20)
  assert.strictEqual(result.cells[1][0].cellType,'blank')
  assert.deepStrictEqual(result.summary,{formulas:3,constants:2,blanks:1,unknown:0})
})

test('single-cell formula is normalized to a 2D result', () => {
  const sheet = sheetFor('B7','=A7+1',42,'42')
  const r = withApi({GetSheet:()=>sheet}, () => inspector.formulaInspectorCommand('Data','B7',26000))
  assert.strictEqual(r.cells[0][0].address,'B7')
  assert.strictEqual(r.cells[0][0].formulaStatus,'formula')
  assert.strictEqual(r.cells[0][0].formula,'=A7+1')
  assert.strictEqual(r.cells[0][0].calculatedValue,42)
})

test('unsupported GetFormula fails closed to unknown instead of guessing formula state', () => {
  const sheet = {GetRange:()=>({GetValue:()=>[[2,4]],GetText:()=>[['2','4']]})}
  const r = withApi({GetSheet:()=>sheet}, () => inspector.formulaInspectorCommand('Data','A1:B1',26000))
  assert.strictEqual(r.ok,true)
  assert.strictEqual(r.cells[0][0].formulaStatus,'unknown')
  assert.strictEqual(r.cells[0][0].formula,null)
  assert.strictEqual(r.cells[0][0].cellType,null)
  assert.strictEqual(r.summary.unknown,2)
  assert(r.unsupported.some((x)=>x.field==='GetFormula'))
})

test('missing calculated-value getter is explicit null/unsupported', () => {
  const sheet = {GetRange:()=>({GetFormula:()=>[['=1+1']],GetText:()=>[['2']]})}
  const r = withApi({GetSheet:()=>sheet}, () => inspector.formulaInspectorCommand('Data','A1',26000))
  assert.strictEqual(r.cells[0][0].formulaStatus,'formula')
  assert.strictEqual(r.cells[0][0].calculatedValue,null)
  assert(r.unsupported.some((x)=>x.field==='GetValue'))
})

test('large bulk range is handled in one range read and respects the limit', () => {
  const rows=1000, cols=26
  const matrix=(fn)=>Array.from({length:rows},(_,r)=>Array.from({length:cols},(_,c)=>fn(r,c)))
  const formulas=matrix(()=>''), values=matrix((r,c)=>r*cols+c), texts=matrix((r,c)=>String(r*cols+c))
  let calls=0
  const sheet={GetRange(address){calls+=1; assert.strictEqual(address,'A1:Z1000'); return {GetFormula:()=>formulas,GetValue:()=>values,GetText:()=>texts}}}
  const r=withApi({GetSheet:()=>sheet},()=>inspector.formulaInspectorCommand('Big','A1:Z1000',26000))
  assert.strictEqual(r.ok,true); assert.strictEqual(r.cellCount,26000); assert.strictEqual(calls,1); assert.strictEqual(r.cells[999][25].address,'Z1000')
})

test('range above the limit fails before touching worksheet range', () => {
  let touched=false
  const r=withApi({GetSheet:()=>({GetRange(){touched=true;return null}})},()=>inspector.formulaInspectorCommand('Big','A1:Z1001',26000))
  assert.strictEqual(r.ok,false); assert.strictEqual(r.outcome,'range-too-large'); assert.strictEqual(touched,false)
})

test('M3.1 wrapper preserves live co-edit boundary and has no file-format fallback', () => {
  const entry=fs.readFileSync('./euro-mcp-m31.cjs','utf8')
  const implementation=fs.readFileSync('./formula-inspector.cjs','utf8')
  assert.match(entry,/office_inspect_formulas/)
  assert.match(entry,/require\('\.\/euro-mcp-m2\.cjs'\)/)
  assert.match(implementation,/callCommand/)
  assert.match(implementation,/spreadsheeteditor/)
  for (const src of [entry,implementation]) {
    assert.doesNotMatch(src,/\brequire\([^\n]*(?:runner|box-helper|adm-zip|xlsx|xml)/i)
    assert.doesNotMatch(src,/\brunJob\s*\(|\bexecFileSync\s*\(/)
  }
})

console.log(`${passed} formula inspector tests passed`)
