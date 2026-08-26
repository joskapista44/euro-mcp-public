'use strict'

const rangeReader = require('./range-reader.cjs')
const formulaInspector = require('./formula-inspector.cjs')

const LIVE_SOURCE = 'live-coedit-editor'
const ERROR_TYPES = Object.freeze({
  '#NULL!':'NULL',
  '#DIV/0!':'DIV/0',
  '#VALUE!':'VALUE',
  '#REF!':'REF',
  '#NAME?':'NAME',
  '#NUM!':'NUM',
  '#N/A':'N/A',
  '#SPILL!':'SPILL',
  '#CALC!':'CALC',
  '#GETTING_DATA':'GETTING_DATA'
})

function normalizeErrorToken(value) {
  if (typeof value !== 'string') return null
  const token = value.trim().toUpperCase()
  return Object.prototype.hasOwnProperty.call(ERROR_TYPES, token) ? token : null
}

function unsupportedField(inspected, name) {
  return Array.isArray(inspected && inspected.unsupported) && inspected.unsupported.some((x) => x && x.field === name)
}

function classifyFormulaError(cell, inspected) {
  if (!cell || cell.formulaStatus !== 'formula') {
    return { errorStatus:'not-formula', errorType:null, errorToken:null, evidence:null }
  }

  const valueToken = normalizeErrorToken(cell.calculatedValue)
  const textToken = normalizeErrorToken(cell.displayText)
  const token = textToken || valueToken
  if (token) {
    const sources = []
    if (valueToken === token) sources.push('GetValue')
    if (textToken === token) sources.push('GetText')
    return {
      errorStatus:'error',
      errorType:ERROR_TYPES[token],
      errorToken:token,
      evidence:{sources, calculatedValue:cell.calculatedValue, displayText:cell.displayText}
    }
  }

  const valueUnavailable = unsupportedField(inspected, 'GetValue')
  const textUnavailable = unsupportedField(inspected, 'GetText')
  if (valueUnavailable && textUnavailable) {
    return {errorStatus:'unknown',errorType:null,errorToken:null,evidence:null,reason:'GetValue and GetText are unavailable; formula error status cannot be proven'}
  }

  return { errorStatus:'ok', errorType:null, errorToken:null, evidence:null }
}

function diagnoseInspection(inspected) {
  if (!inspected || inspected.ok !== true) return inspected
  if (!Array.isArray(inspected.cells)) {
    return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector did not return a cell matrix'}
  }

  const cells = []
  const errors = []
  const summary = {cells:0,formulas:0,formulaErrors:0,formulaOk:0,notFormula:0,unknown:0,byType:{}}

  for (const sourceRow of inspected.cells) {
    if (!Array.isArray(sourceRow)) return {ok:false,outcome:'unsupported-result',source:LIVE_SOURCE,error:'formula inspector returned a non-rectangular cell matrix'}
    const row = []
    for (const cell of sourceRow) {
      const diagnosis = classifyFormulaError(cell, inspected)
      const item = {
        address:cell.address,
        row:cell.row,
        column:cell.column,
        formulaStatus:cell.formulaStatus,
        formula:cell.formula,
        calculatedValue:cell.calculatedValue,
        displayText:cell.displayText,
        cellType:cell.cellType,
        ...diagnosis
      }
      row.push(item)
      summary.cells += 1
      if (cell.formulaStatus === 'formula') summary.formulas += 1
      if (diagnosis.errorStatus === 'error') {
        summary.formulaErrors += 1
        summary.byType[diagnosis.errorType] = (summary.byType[diagnosis.errorType] || 0) + 1
        errors.push({address:cell.address,errorType:diagnosis.errorType,errorToken:diagnosis.errorToken,formula:cell.formula,evidence:diagnosis.evidence})
      } else if (diagnosis.errorStatus === 'ok') summary.formulaOk += 1
      else if (diagnosis.errorStatus === 'not-formula') summary.notFormula += 1
      else summary.unknown += 1
    }
    cells.push(row)
  }

  return {
    ok:true,
    outcome:'ok',
    source:LIVE_SOURCE,
    sheet:inspected.sheet,
    range:inspected.range,
    rows:inspected.rows,
    columns:inspected.columns,
    cellCount:inspected.cellCount,
    cells,
    errors,
    summary,
    unsupported:inspected.unsupported || []
  }
}

async function diagnoseFormulaErrorsInFrame(frame, apiHely, {sheet, range, maxCells=26000, callbackTimeoutMs=15000}) {
  const parsed = rangeReader.parseA1Range(range)
  if (!parsed) return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if (parsed.cellCount > maxCells) return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const inspected = await formulaInspector.inspectFormulaInFrame(frame, apiHely, {sheet,range:parsed.address,maxCells,callbackTimeoutMs})
  return diagnoseInspection(inspected)
}

async function diagnoseFormulaErrorsLive({url,user,pass,fileId,sheet,range,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}) {
  const parsed = rangeReader.parseA1Range(range)
  if (!parsed) return {ok:false,outcome:'invalid-range',source:LIVE_SOURCE,sheet,range,error:'range must be a rectangular A1 address such as A1:H50'}
  if (parsed.cellCount > maxCells) return {ok:false,outcome:'range-too-large',source:LIVE_SOURCE,sheet,range:parsed.address,cellCount:parsed.cellCount,maxCells,error:'requested range exceeds the configured cell limit'}
  const loaded = loadPlaywright()
  if (!loaded.ok) return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium} = loaded.pw
  const browser = await chromium.launch()
  try {
    const ctx = await browser.newContext({viewport:{width:1400,height:900}})
    const page = await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs})
    await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')])
    await page.waitForTimeout(2500)
    if (/\/login/.test(page.url())) return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed (browser remained on login page)'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs})
    await page.waitForTimeout(22000)
    const frame = page.frames().find((f) => /spreadsheeteditor/.test(f.url()))
    if (!frame) return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely = await frame.evaluate(() => {
      if ((window.Asc || {}).editor && typeof window.Asc.editor.callCommand === 'function') return 'window.Asc.editor'
      if (window.editor && typeof window.editor.callCommand === 'function') return 'window.editor'
      return null
    })
    if (!apiHely) return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const result = await diagnoseFormulaErrorsInFrame(frame,apiHely,{sheet,range:parsed.address,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

module.exports={ERROR_TYPES,normalizeErrorToken,classifyFormulaError,diagnoseInspection,diagnoseFormulaErrorsInFrame,diagnoseFormulaErrorsLive}
