'use strict'
const assert=require('assert')
const ff=require('./formula-fill.cjs')

function eq(actual,expected){assert.strictEqual(actual,expected)}

let t=ff.translateFormulaA1('=A1+$A$1+$A1+A$1',1,2); eq(t.ok,true); eq(t.formula,'=C2+$A$1+$A2+C$1')
t=ff.translateFormulaA1('=SUM(A1:B2)+"A1"',2,1); eq(t.ok,true); eq(t.formula,'=SUM(B3:C4)+"A1"')
t=ff.translateFormulaA1("='Other Sheet'!A1+$B2+C$3",3,2); eq(t.ok,true); eq(t.formula,"='Other Sheet'!C4+$B5+E$3")
t=ff.translateFormulaA1('=LOG10(A1)',1,1); eq(t.ok,true); eq(t.formula,'=LOG10(B2)')
t=ff.translateFormulaA1('=A1',-1,0); eq(t.outcome,'reference-out-of-bounds')

let p=ff.buildFillMatrix('B2','B2:B5','=A2*$C$1'); eq(p.ok,true); eq(p.direction,'vertical'); assert.deepStrictEqual(p.formulas,[['=A2*$C$1'],['=A3*$C$1'],['=A4*$C$1'],['=A5*$C$1']])
p=ff.buildFillMatrix('B2','B2:E2','=A2+$A2+A$2+$A$2'); eq(p.ok,true); eq(p.direction,'horizontal'); assert.deepStrictEqual(p.formulas,[['=A2+$A2+A$2+$A$2','=B2+$A2+B$2+$A$2','=C2+$A2+C$2+$A$2','=D2+$A2+D$2+$A$2']])
eq(ff.buildFillMatrix('B2','B2:C3','=A1').outcome,'unsupported-shape')
eq(ff.buildFillMatrix('bad','B2:B3','=A1').outcome,'invalid-source-cell')
eq(ff.buildFillMatrix('B2','A1:Z1001','=A1').outcome,'range-too-large')

const cells={B2:'=A2+$A$1',B3:null,B4:null}
global.Api={GetSheet:(name)=>name==='Second'?{GetRange:(addr)=>({
  GetFormula:()=>addr.includes(':')?[[cells.B2],[cells.B3],[cells.B4]]:cells[addr],
  GetValue:()=>addr.includes(':')?[[1],[2],[3]]:1,
  GetText:()=>addr.includes(':')?[['1'],['2'],['3']]:'1',
  SetFormula:(f)=>{cells[addr]=f}
})}:null}
const frame={evaluate:async(fn,args)=>{global.window={Asc:{editor:{callCommand:(cmd,a,b,cb)=>cb(cmd())}}}; return fn(args)}}
;(async()=>{
  const r=await ff.fillFormulaInFrame(frame,'window.Asc.editor',{sheet:'Second',sourceCell:'B2',targetRange:'B2:B4',callbackTimeoutMs:50})
  eq(r.ok,true); eq(r.verified,true); eq(r.direction,'vertical'); eq(cells.B3,'=A3+$A$1'); eq(cells.B4,'=A4+$A$1')
  console.log('ok - M3.3 formula fill tests')
})().catch(e=>{console.error(e);process.exit(1)})
