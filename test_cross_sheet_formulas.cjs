'use strict'
const assert=require('assert')
const xs=require('./cross-sheet-formulas.cjs')

function eq(a,b){assert.strictEqual(a,b)}

let q=xs.qualifyA1('Sales Data','A1'); eq(q.ok,true); eq(q.formulaRef,"'Sales Data'!A1")
q=xs.qualifyA1("O'Brien",'$A$1:B$5'); eq(q.ok,true); eq(q.formulaRef,"'O''Brien'!$A$1:B$5")
eq(xs.qualifyA1('Sheet1','bad').outcome,'invalid-reference')

let s=xs.scanSheetReferences("='Sales Data'!A1+Sheet2!$B$2+'O''Brien'!C$3"); eq(s.ok,true); assert.deepStrictEqual(s.sheets,['Sales Data','Sheet2',"O'Brien"])
s=xs.scanSheetReferences('="Sheet1!A1"+Sheet2!A1'); eq(s.ok,true); assert.deepStrictEqual(s.sheets,['Sheet2'])
eq(xs.scanSheetReferences("='Broken!A1").outcome,'invalid-sheet-reference')

let v=xs.validateCrossSheetMatrix([["=Sheet1!A1+$B2+C$3", "='Sales Data'!A1:B2"]]); eq(v.ok,true); assert.deepStrictEqual(v.sheets,['Sheet1','Sales Data'])
eq(xs.validateCrossSheetMatrix([['=A1+1']]).outcome,'missing-cross-sheet-reference')
eq(xs.validateCrossSheetMatrix([["='Bad!A1"]]).outcome,'invalid-sheet-reference')

const formulas={C2:"='Source Data'!A1+$B$2",C3:null}
global.Api={GetSheet:(name)=>{
  if(name==='Target')return {GetRange:(addr)=>({SetFormula:(f)=>{formulas[addr]=f},GetFormula:()=>addr.includes(':')?[[formulas.C2],[formulas.C3]]:formulas[addr],GetValue:()=>1,GetText:()=>String(1)})}
  if(name==='Source Data')return {GetRange:()=>({})}
  return null
}}
const frame={evaluate:async(fn,args)=>{global.window={Asc:{editor:{callCommand:(cmd,a,b,cb)=>cb(cmd())}}};return fn(args)}}

;(async()=>{
  let p=await xs.preflightReferencedSheets(frame,'window.Asc.editor',['Source Data']); eq(p.ok,true)
  p=await xs.preflightReferencedSheets(frame,'window.Asc.editor',['Missing']); eq(p.outcome,'referenced-sheet-not-found')
  const r=await xs.writeCrossSheetInFrame(frame,'window.Asc.editor',{sheet:'Target',range:'C2:C3',formulas:[["='Source Data'!A1+$B$2"],["='Source Data'!A2+$B$2"]],callbackTimeoutMs:1000})
  eq(r.ok,true); eq(r.verified,true); assert.deepStrictEqual(r.referencedSheets,['Source Data'])
  const missing=await xs.writeCrossSheetInFrame(frame,'window.Asc.editor',{sheet:'Target',range:'C2',formulas:[["=Missing!A1"]],callbackTimeoutMs:1000}); eq(missing.outcome,'referenced-sheet-not-found')

  const fs=require('fs')
  const prod=fs.readFileSync('cross-sheet-formulas.cjs','utf8')+fs.readFileSync('euro-mcp-m34.cjs','utf8')
  for(const forbidden of ['DocBuilder','soffice','adm-zip','xlsx.write','zip/xml']) assert.ok(!prod.toLowerCase().includes(forbidden.toLowerCase()),`forbidden fallback token: ${forbidden}`)
  console.log('test_cross_sheet_formulas: PASS')
})().catch(e=>{console.error(e);process.exit(1)})
