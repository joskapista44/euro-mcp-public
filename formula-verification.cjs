'use strict'

const rangeReader = require('./range-reader.cjs')
const formulaInspector = require('./formula-inspector.cjs')
const formulaErrors = require('./formula-error-intelligence.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function normalizeExpected(expected, rows, columns) {
  if (expected === null || expected === undefined) return { ok:true, matrix:null }
  if (!Array.isArray(expected) || expected.length !== rows) return { ok:false, error:'expectedFormulas must be a 2D array matching the requested range shape' }
  const matrix = []
  for (let r=0;r<rows;r++) {
    if (!Array.isArray(expected[r]) || expected[r].length !== columns) return { ok:false, error:'expectedFormulas must be a rectangular 2D array matching the requested range shape' }
    matrix.push(expected[r].map((v)=>v === null || v === undefined || v === '' ? null : String(v)))
  }
  return { ok:true, matrix }
}

function verifyInspection(inspected, expectedFormulas = null) {
  if (!inspected || inspected.ok !== true) return inspected || {ok:false,outcome:'invalid-result',source:LIVE_SOURCE,error:'formula inspector returned no structured result'}
  if (inspected.source !== LIVE_SOURCE || inspected.outcome !== 'ok') return {ok:false,outcome:'non-live-or-non-ok',source:inspected.source || null,error:'formula inspection did not prove a successful live co-edit result',cause:inspected}
  if (!Array.isArray(inspected.cells) || !Number.isInteger(inspected.rows) || !Number.isInteger(inspected.columns)) return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector returned an unsupported result shape'}

  const exp = normalizeExpected(expectedFormulas, inspected.rows, inspected.columns)
  if (!exp.ok) return {ok:false,outcome:'invalid-expectation',source:LIVE_SOURCE,error:exp.error}

  const diagnosed = formulaErrors.diagnoseInspection(inspected)
  if (!diagnosed || diagnosed.ok !== true) return diagnosed || {ok:false,outcome:'diagnosis-failed',source:LIVE_SOURCE,error:'formula error diagnosis failed'}

  const cells=[]
  const mismatches=[]
  const unknown=[]
  const errors=[]
  const summary={cells:0,matched:0,mismatched:0,unchecked:0,formulaErrors:0,unknown:0,formulaOk:0,constants:0}

  for (let r=0;r<inspected.rows;r++) {
    if (!Array.isArray(inspected.cells[r]) || inspected.cells[r].length !== inspected.columns) return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector returned a non-rectangular cell matrix'}
    const row=[]
    for (let c=0;c<inspected.columns;c++) {
      const raw=inspected.cells[r][c]
      const diag=diagnosed.cells[r][c]
      const expected=exp.matrix ? exp.matrix[r][c] : undefined
      let matchStatus='unchecked'
      let reason=null
      if (exp.matrix) {
        if (raw.formulaStatus === 'unknown') { matchStatus='unknown'; reason='formula status is not proven' }
        else if (expected === null) {
          matchStatus = raw.formulaStatus === 'formula' ? 'mismatch' : 'match'
          if (matchStatus === 'mismatch') reason='expected no formula but live read-back contains a formula'
        } else if (raw.formulaStatus !== 'formula') {
          matchStatus='mismatch'; reason='expected a formula but live read-back does not prove a formula'
        } else if (raw.formula === expected) matchStatus='match'
        else { matchStatus='mismatch'; reason='live formula text differs from expected formula text' }
      }

      const item={address:raw.address,row:raw.row,column:raw.column,expectedFormula:expected === undefined ? null : expected,actualFormula:raw.formula || null,formulaStatus:raw.formulaStatus,formulaMatch:matchStatus,calculatedValue:raw.calculatedValue,displayText:raw.displayText,errorStatus:diag.errorStatus,errorType:diag.errorType,errorToken:diag.errorToken,evidence:diag.evidence || null,reason:reason || diag.reason || null}
      row.push(item); summary.cells++
      if (matchStatus==='match') summary.matched++
      else if (matchStatus==='mismatch') { summary.mismatched++; mismatches.push(item) }
      else if (matchStatus==='unknown') { summary.unknown++; unknown.push(item) }
      else summary.unchecked++
      if (diag.errorStatus==='error') { summary.formulaErrors++; errors.push(item) }
      else if (diag.errorStatus==='unknown') { if (matchStatus!=='unknown') summary.unknown++; unknown.push(item) }
      else if (diag.errorStatus==='ok') summary.formulaOk++
      else if (diag.errorStatus==='not-formula') summary.constants++
    }
    cells.push(row)
  }

  let verificationStatus='verified'
  if (summary.mismatched>0) verificationStatus='mismatch'
  else if (summary.unknown>0) verificationStatus='unknown'
  else if (!exp.matrix) verificationStatus='observed'

  return {ok:true,outcome:'ok',source:LIVE_SOURCE,verificationStatus,sheet:inspected.sheet,range:inspected.range,rows:inspected.rows,columns:inspected.columns,cellCount:inspected.cellCount,cells,mismatches,unknown,errors,summary,unsupported:inspected.unsupported || []}
}

async function verifyFormulaInFrame(frame, apiHely, {sheet,range,expectedFormulas=null,maxCells=26000,callbackTimeoutMs=15000}) {
  const parsed=rangeReader.parseA1Range(range)
  if (!parsed) return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if (parsed.cellCount>maxCells) return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const inspected=await formulaInspector.inspectFormulaInFrame(frame,apiHely,{sheet,range:parsed.address,maxCells,callbackTimeoutMs})
  return verifyInspection(inspected,expectedFormulas)
}

async function verifyFormulaLive({url,user,pass,fileId,sheet,range,expectedFormulas=null,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}) {
  const parsed=rangeReader.parseA1Range(range)
  if (!parsed) return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if (parsed.cellCount>maxCells) return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const loaded=loadPlaywright()
  if (!loaded.ok) return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium}=loaded.pw
  const browser=await chromium.launch()
  try {
    const ctx=await browser.newContext({viewport:{width:1400,height:900}})
    const page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs})
    await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed (browser remained on login page)'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs})
    await page.waitForTimeout(22000)
    const frame=page.frames().find((f)=>/spreadsheeteditor/.test(f.url()))
    if (!frame) return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>{if((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')return 'window.Asc.editor';if(window.editor&&typeof window.editor.callCommand==='function')return 'window.editor';return null})
    if (!apiHely) return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const result=await verifyFormulaInFrame(frame,apiHely,{sheet,range:parsed.address,expectedFormulas,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

module.exports={normalizeExpected,verifyInspection,verifyFormulaInFrame,verifyFormulaLive}
