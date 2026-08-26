'use strict'

const rangeReader = require('./range-reader.cjs')
const formulaInspector = require('./formula-inspector.cjs')
const formulaErrors = require('./formula-error-intelligence.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function validateExpectedFormulaMatrix(parsed, expectedFormulas) {
  if (expectedFormulas == null) return {ok:true,expected:false}
  if (!Array.isArray(expectedFormulas) || expectedFormulas.length !== parsed.rows) {
    return {ok:false,outcome:'dimension-mismatch',error:'expected_formulas rows must exactly match target range'}
  }
  for (let r=0;r<expectedFormulas.length;r++) {
    const row=expectedFormulas[r]
    if (!Array.isArray(row) || row.length !== parsed.columns) {
      return {ok:false,outcome:'dimension-mismatch',error:'expected_formulas columns must exactly match target range',row:r+1}
    }
    for (let c=0;c<row.length;c++) {
      if (typeof row[c] !== 'string' || !row[c].startsWith('=')) {
        return {ok:false,outcome:'invalid-expected-formula',error:'every expected formula must be a string beginning with =',row:r+1,column:c+1}
      }
    }
  }
  return {ok:true,expected:true}
}

function verifyFormulaCell(cell, expectedFormula, diagnosis) {
  let formulaMatch='not-requested'
  let matchReason=null
  if (expectedFormula !== undefined) {
    if (!cell || cell.formulaStatus === 'unknown' || cell.formulaStatus == null) {
      formulaMatch='unknown'; matchReason='formula read-back status is not proven'
    } else if (cell.formulaStatus !== 'formula') {
      formulaMatch='mismatch'; matchReason='target cell is not a formula after the operation'
    } else if (typeof cell.formula !== 'string') {
      formulaMatch='unknown'; matchReason='formula text is unavailable from live read-back'
    } else if (cell.formula === expectedFormula) {
      formulaMatch='match'
    } else {
      formulaMatch='mismatch'; matchReason='formula read-back differs from expected formula'
    }
  }

  let verificationStatus='verified'
  if (formulaMatch === 'unknown' || (diagnosis && diagnosis.errorStatus === 'unknown')) verificationStatus='unknown'
  else if (formulaMatch === 'mismatch' || (diagnosis && diagnosis.errorStatus === 'error')) verificationStatus='failed'
  else if (formulaMatch === 'not-requested') verificationStatus=(diagnosis && diagnosis.errorStatus === 'ok')?'observed-ok':'observed'

  return {formulaMatch,matchReason,verificationStatus}
}

function verifyInspection(inspected, expectedFormulas) {
  if (!inspected || inspected.ok !== true) return inspected
  if (!Array.isArray(inspected.cells)) {
    return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector did not return a cell matrix'}
  }

  const diagnosed=formulaErrors.diagnoseInspection(inspected)
  if (!diagnosed || diagnosed.ok !== true || !Array.isArray(diagnosed.cells)) return diagnosed

  if (expectedFormulas != null) {
    if (!Array.isArray(expectedFormulas) || expectedFormulas.length !== inspected.cells.length) {
      return {ok:false,outcome:'dimension-mismatch',source:LIVE_SOURCE,error:'expected_formulas rows must exactly match inspected range'}
    }
  }

  const cells=[]
  const mismatches=[]
  const unknown=[]
  const formulaErrorsFound=[]
  const summary={cells:0,expected:expectedFormulas!=null,matched:0,mismatched:0,matchUnknown:0,formulaErrors:0,errorUnknown:0,verified:0,failed:0,unknown:0,observed:0}

  for(let r=0;r<inspected.cells.length;r++){
    const sourceRow=inspected.cells[r]
    const diagRow=diagnosed.cells[r]
    if(!Array.isArray(sourceRow)||!Array.isArray(diagRow)) return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector returned a non-rectangular cell matrix'}
    if(expectedFormulas!=null && (!Array.isArray(expectedFormulas[r]) || expectedFormulas[r].length!==sourceRow.length)) {
      return {ok:false,outcome:'dimension-mismatch',source:LIVE_SOURCE,error:'expected_formulas columns must exactly match inspected range',row:r+1}
    }
    const row=[]
    for(let c=0;c<sourceRow.length;c++){
      const cell=sourceRow[c]||{}
      const diagnosis=diagRow[c]||{}
      const expected=expectedFormulas!=null?expectedFormulas[r][c]:undefined
      const verified=verifyFormulaCell(cell,expected,diagnosis)
      const item={
        address:cell.address,row:cell.row,column:cell.column,
        expectedFormula:expected===undefined?null:expected,
        formulaStatus:cell.formulaStatus,formula:cell.formula,
        calculatedValue:cell.calculatedValue,displayText:cell.displayText,cellType:cell.cellType,
        formulaMatch:verified.formulaMatch,matchReason:verified.matchReason,
        errorStatus:diagnosis.errorStatus,errorType:diagnosis.errorType,errorToken:diagnosis.errorToken,errorEvidence:diagnosis.evidence||null,errorReason:diagnosis.reason||null,
        verificationStatus:verified.verificationStatus
      }
      row.push(item); summary.cells++
      if(verified.formulaMatch==='match')summary.matched++
      else if(verified.formulaMatch==='mismatch'){summary.mismatched++;mismatches.push({address:item.address,expectedFormula:item.expectedFormula,actualFormula:item.formula,formulaStatus:item.formulaStatus,reason:item.matchReason})}
      else if(verified.formulaMatch==='unknown')summary.matchUnknown++
      if(diagnosis.errorStatus==='error'){summary.formulaErrors++;formulaErrorsFound.push({address:item.address,errorType:item.errorType,errorToken:item.errorToken,formula:item.formula})}
      else if(diagnosis.errorStatus==='unknown')summary.errorUnknown++
      if(verified.verificationStatus==='verified')summary.verified++
      else if(verified.verificationStatus==='failed')summary.failed++
      else if(verified.verificationStatus==='unknown'){summary.unknown++;unknown.push({address:item.address,formulaMatch:item.formulaMatch,errorStatus:item.errorStatus,matchReason:item.matchReason,errorReason:item.errorReason})}
      else summary.observed++
    }
    cells.push(row)
  }

  let verificationOutcome='observed'
  if(expectedFormulas!=null){
    if(summary.unknown>0)verificationOutcome='unknown'
    else if(summary.failed>0)verificationOutcome='failed'
    else verificationOutcome='verified'
  }

  return {
    ok:true,outcome:'ok',source:LIVE_SOURCE,
    sheet:inspected.sheet,range:inspected.range,rows:inspected.rows,columns:inspected.columns,cellCount:inspected.cellCount,
    verificationOutcome,summary,cells,mismatches,unknown,formulaErrors:formulaErrorsFound,
    unsupported:inspected.unsupported||[]
  }
}

async function verifyFormulaRangeInFrame(frame, apiHely, {sheet,range,expectedFormulas=null,maxCells=26000,callbackTimeoutMs=15000}) {
  const parsed=rangeReader.parseA1Range(range)
  if(!parsed)return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if(parsed.cellCount>maxCells)return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const valid=validateExpectedFormulaMatrix(parsed,expectedFormulas)
  if(!valid.ok)return {...valid,source:LIVE_SOURCE,sheet,range:parsed.address}
  const inspected=await formulaInspector.inspectFormulaInFrame(frame,apiHely,{sheet,range:parsed.address,maxCells,callbackTimeoutMs})
  return verifyInspection(inspected,expectedFormulas)
}

async function verifyFormulaRangeLive({url,user,pass,fileId,sheet,range,expectedFormulas=null,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}) {
  const parsed=rangeReader.parseA1Range(range)
  if(!parsed)return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if(parsed.cellCount>maxCells)return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const valid=validateExpectedFormulaMatrix(parsed,expectedFormulas)
  if(!valid.ok)return {...valid,source:LIVE_SOURCE,sheet,range:parsed.address}
  const loaded=loadPlaywright()
  if(!loaded.ok)return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium}=loaded.pw
  const browser=await chromium.launch()
  try{
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}); const page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]); await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed (browser remained on login page)'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.waitForTimeout(22000)
    const frame=page.frames().find((f)=>/spreadsheeteditor/.test(f.url())); if(!frame)return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>{if((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')return 'window.Asc.editor';if(window.editor&&typeof window.editor.callCommand==='function')return 'window.editor';return null})
    if(!apiHely)return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const result=await verifyFormulaRangeInFrame(frame,apiHely,{sheet,range:parsed.address,expectedFormulas,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally {await browser.close().catch(()=>{})}
}

module.exports={validateExpectedFormulaMatrix,verifyFormulaCell,verifyInspection,verifyFormulaRangeInFrame,verifyFormulaRangeLive}
