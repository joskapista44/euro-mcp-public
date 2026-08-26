'use strict'

const rangeReader = require('./range-reader.cjs')
const formulaWriter = require('./formula-writer.cjs')

const LIVE_SOURCE='live-coedit-editor'

function colNumber(label){ let n=0; for(const ch of String(label).toUpperCase()) n=n*26+ch.charCodeAt(0)-64; return n }
function colLabel(n){ let s=''; while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)} return s }
function parseCell(address){ const m=String(address||'').match(/^([A-Za-z]{1,3})(\d+)$/); if(!m)return null; const column=colNumber(m[1]),row=Number(m[2]); if(column<1||column>16384||row<1||row>1048576)return null; return {address:`${m[1].toUpperCase()}${row}`,row,column} }

function translateFormulaA1(formula,rowDelta,columnDelta){
  if(typeof formula!=='string'||formula[0]!=='=') return {ok:false,outcome:'invalid-formula',error:'source formula must begin with ='}
  let out='',i=0,inString=false
  while(i<formula.length){
    const ch=formula[i]
    if(ch==='"'){ out+=ch; if(inString&&formula[i+1]==='"'){out+='"';i+=2;continue} inString=!inString;i++;continue }
    if(inString){out+=ch;i++;continue}
    const rest=formula.slice(i)
    const m=rest.match(/^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)/)
    if(m){
      const before=i===0?'':formula[i-1]
      const after=formula[i+m[0].length]||''
      const c=colNumber(m[3]),r=Number(m[5])
      const boundaryOk=!/[A-Za-z0-9_.]/.test(before)&&after!=='('&&!/[A-Za-z0-9_]/.test(after)
      if(boundaryOk&&c>=1&&c<=16384&&r>=1&&r<=1048576){
        const nc=m[2]==='$'?c:c+columnDelta, nr=m[4]==='$'?r:r+rowDelta
        if(nc<1||nc>16384||nr<1||nr>1048576)return {ok:false,outcome:'reference-out-of-bounds',error:'relative fill would move a reference outside worksheet bounds'}
        out+=(m[1]||'')+m[2]+colLabel(nc)+m[4]+nr; i+=m[0].length; continue
      }
    }
    out+=ch;i++
  }
  if(inString)return {ok:false,outcome:'unsupported-formula',error:'unterminated string literal in source formula'}
  return {ok:true,formula:out}
}

function buildFillMatrix(sourceCell,targetRange,sourceFormula,maxCells=26000){
  const src=parseCell(sourceCell), target=rangeReader.parseA1Range(targetRange)
  if(!src)return {ok:false,outcome:'invalid-source-cell',error:'source_cell must be one A1 cell'}
  if(!target)return {ok:false,outcome:'invalid-range',error:'target_range must be rectangular A1'}
  if(target.cellCount>maxCells)return {ok:false,outcome:'range-too-large',cellCount:target.cellCount,maxCells,error:'target range exceeds configured cell limit'}
  if(target.rows>1&&target.columns>1)return {ok:false,outcome:'unsupported-shape',error:'M3.3 fill supports one-dimensional vertical or horizontal target ranges'}
  const matrix=[]
  for(let r=0;r<target.rows;r++){
    const row=[]
    for(let c=0;c<target.columns;c++){
      const tr=target.start.row+r,tc=target.start.column+c
      const translated=translateFormulaA1(sourceFormula,tr-src.row,tc-src.column)
      if(!translated.ok)return {...translated,targetCell:`${colLabel(tc)}${tr}`}
      row.push(translated.formula)
    }
    matrix.push(row)
  }
  return {ok:true,source:src,target,formulas:matrix,direction:target.rows>1?'vertical':target.columns>1?'horizontal':'single'}
}

async function fillFormulaInFrame(frame,apiHely,{sheet,sourceCell,targetRange,maxCells=26000,callbackTimeoutMs=15000}){
  const src=parseCell(sourceCell); if(!src)return {ok:false,outcome:'invalid-source-cell',source:LIVE_SOURCE,error:'source_cell must be one A1 cell'}
  const inspected=await rangeReader.readRangeInFrame(frame,apiHely,{sheet,range:src.address,maxCells:1,callbackTimeoutMs})
  const cell=inspected&&inspected.ok&&inspected.cells&&inspected.cells[0]&&inspected.cells[0][0]
  if(!cell||cell.dataType!=='formula'||typeof cell.formula!=='string')return {ok:false,outcome:'source-not-formula',source:LIVE_SOURCE,sheet,sourceCell:src.address,error:'source cell does not expose a formula through live range reader'}
  const plan=buildFillMatrix(src.address,targetRange,cell.formula,maxCells); if(!plan.ok)return {...plan,source:LIVE_SOURCE,sheet,sourceCell:src.address,targetRange}
  const write=await formulaWriter.writeFormulaInFrame(frame,apiHely,{sheet,range:plan.target.address,formulas:plan.formulas,maxCells,callbackTimeoutMs})
  if(!write.ok)return {...write,sourceCell:src.address,direction:plan.direction}
  const {readBack,verification}=await formulaWriter.verifyFormulaRangeInFrame(frame,apiHely,{sheet,range:plan.target.address,formulas:plan.formulas,maxCells,callbackTimeoutMs})
  return {...write,sourceCell:src.address,direction:plan.direction,expectedFormulas:plan.formulas,verified:verification.ok,verification,readBack}
}

async function fillFormulaLive({url,user,pass,fileId,sheet,sourceCell,targetRange,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}){
  const loaded=loadPlaywright(); if(!loaded.ok)return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium}=loaded.pw,browser=await chromium.launch()
  try{
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]); await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.waitForTimeout(22000)
    const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url())); if(!frame)return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>{if((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')return 'window.Asc.editor';if(window.editor&&typeof window.editor.callCommand==='function')return 'window.editor';return null})
    if(!apiHely)return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const result=await fillFormulaInFrame(frame,apiHely,{sheet,sourceCell,targetRange,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally {await browser.close().catch(()=>{})}
}

module.exports={parseCell,translateFormulaA1,buildFillMatrix,fillFormulaInFrame,fillFormulaLive}
