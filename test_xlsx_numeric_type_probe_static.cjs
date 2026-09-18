'use strict'
const assert=require('assert/strict')
const {probeCommand}=require('./test_xlsx_numeric_type_live_probe.cjs')
const previous=global.Api
try {
  const range={GetValue:()=> '120',GetValue2:()=> '120',GetText:()=> '120',GetFormula:()=> '120',GetNumberFormat:()=> 'General'}
  // Reproduce runtime: every TYPE call returns error code 16. Never call this PASS.
  global.Api={GetSheet:()=>({GetRange:()=>range}),WorksheetFunction:{TYPE:()=>16}}
  let result=probeCommand('Data')
  assert.equal(result.typeControlsPass,false)
  assert.equal(result.outcome,'diagnostic-collected-not-acceptance')
  assert.equal(result.cells.length,6)
  assert.equal(result.cells[3].getters.GetValue.value,'120')
  // Literal controls pass but range conversion fails: keep both pieces of evidence.
  global.Api.WorksheetFunction.TYPE=value=>typeof value==='number'?1:typeof value==='string'?2:typeof value==='boolean'?4:16
  result=probeCommand('Data')
  assert.equal(result.typeControlsPass,true)
  assert.equal(result.cells[3].reference.TYPE.value,16)
  assert.equal(result.cells[3].valueArgument.TYPE.value,2)
  assert.equal(result.cells[3].reference.ISNUMBER.error,'unavailable')
  console.log('XLSX NUMERIC TYPE PROBE STATIC: PASS')
} finally { if(previous===undefined)delete global.Api;else global.Api=previous }
