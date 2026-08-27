'use strict'

const assert = require('assert')
const fs = require('fs')
const path = require('path')
const { resolveNumberFormat, numberFormatCommand } = require('./number-formatting.cjs')

function makeRange({ value=1234.5, text='$1,234.50', format='General', exposeFormat=true, exposeValue=true, exposeText=true }={}) {
  let currentFormat=format
  const r={
    SetNumberFormat(v){ currentFormat=v; return true }
  }
  if(exposeFormat) r.GetNumberFormat=()=>currentFormat
  if(exposeValue) r.GetValue2=()=>value
  if(exposeText) r.GetText=()=>text
  return r
}

function withApi(range, fn) {
  const old=global.Api
  global.Api={ GetSheet(name){ return name==='Sheet1'?{ GetRange(){ return range } }:null } }
  try { return fn() } finally { global.Api=old }
}

// preset resolution
assert.deepStrictEqual(resolveNumberFormat({kind:'percent',decimals:1}),{ok:true,format:'0.0%'})
assert.deepStrictEqual(resolveNumberFormat({kind:'decimal',decimals:3}),{ok:true,format:'#,##0.000'})
assert.deepStrictEqual(resolveNumberFormat({kind:'date'}),{ok:true,format:'yyyy-mm-dd'})
assert.deepStrictEqual(resolveNumberFormat({kind:'datetime'}),{ok:true,format:'yyyy-mm-dd hh:mm'})
assert.deepStrictEqual(resolveNumberFormat({kind:'custom',format:'0.0000'}),{ok:true,format:'0.0000'})
assert.deepStrictEqual(resolveNumberFormat({kind:'currency',symbol:'€',decimals:2}),{ok:true,format:'"€"#,##0.00'})
assert.equal(resolveNumberFormat({kind:'percent',decimals:-1}).ok,false)
assert.equal(resolveNumberFormat({kind:'wat'}).ok,false)

// live format + raw/display separation + preservation
{
  const range=makeRange({value:1234.5,text:'€1,234.50'})
  const res=withApi(range,()=>numberFormatCommand('Sheet1','A1','"€"#,##0.00',{expectedValue:1234.5,expectedDisplay:'€1,234.50'}))
  assert.equal(res.ok,true)
  assert.equal(res.source,'live-coedit-editor')
  assert.equal(res.underlyingValue,1234.5)
  assert.equal(res.displayValue,'€1,234.50')
  assert.equal(res.verification.outcome,'pass')
  assert.equal(res.verification.checks.numberFormat.status,'pass')
  assert.equal(res.verification.checks.underlyingValuePreserved.status,'pass')
  assert.equal(res.verification.checks.expectedDisplay.status,'pass')
}

// measurable mismatch => FAIL
{
  const range=makeRange({value:0.15,text:'15.0%'})
  const res=withApi(range,()=>numberFormatCommand('Sheet1','B2','0.0%',{expectedDisplay:'14.0%'}))
  assert.equal(res.ok,false)
  assert.equal(res.outcome,'verification-mismatch')
  assert.equal(res.verification.outcome,'fail')
  assert(res.verification.mismatches.includes('expectedDisplay'))
}

// unavailable getter => UNKNOWN, never guessed PASS
{
  const range=makeRange({exposeFormat:false,exposeText:false})
  const res=withApi(range,()=>numberFormatCommand('Sheet1','C3','0.00',{}))
  assert.equal(res.ok,true)
  assert.equal(res.verification.outcome,'unknown')
  assert(res.verification.unknown.includes('numberFormat'))
  assert(res.verification.unknown.includes('displayValue'))
}

// raw value mutation by formatter is detected fail-closed
{
  let calls=0
  const range={
    SetNumberFormat(){ return true },
    GetNumberFormat(){ return '0.00' },
    GetValue2(){ calls++; return calls===1?10:11 },
    GetText(){ return '11.00' }
  }
  const res=withApi(range,()=>numberFormatCommand('Sheet1','D4','0.00',{}))
  assert.equal(res.ok,false)
  assert(res.verification.mismatches.includes('underlyingValuePreserved'))
}

// architecture guard: specialized M4.2 module must stay live-editor-only
{
  const src=fs.readFileSync(path.join(__dirname,'number-formatting.cjs'),'utf8')
  for(const forbidden of ['soffice','LibreOffice','unzip','xl/worksheets','document.xml']) assert.equal(src.includes(forbidden),false,forbidden)
}

console.log('M4.2 number formatting tests: PASS')
