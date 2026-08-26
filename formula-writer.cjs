'use strict'

const rangeReader = require('./range-reader.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function validateFormulaMatrix(range, formulas, maxCells = 26000) {
  const parsed = rangeReader.parseA1Range(range)
  if (!parsed) return { ok:false, outcome:'invalid-range', error:'range must be a rectangular A1 address such as A1:H50' }
  if (!Array.isArray(formulas) || !formulas.length || !Array.isArray(formulas[0]) || !formulas[0].length) return { ok:false, outcome:'invalid-formulas', error:'formulas must be a non-empty 2D array' }
  const cols = formulas[0].length
  for (let r=0;r<formulas.length;r++) {
    if (!Array.isArray(formulas[r]) || formulas[r].length !== cols) return { ok:false, outcome:'invalid-formulas', error:'formulas must be rectangular' }
    for (let c=0;c<cols;c++) if (typeof formulas[r][c] !== 'string' || !formulas[r][c].startsWith('=')) return { ok:false, outcome:'invalid-formula', error:'every formula cell must be a string beginning with =', row:r+1, column:c+1 }
  }
  if (formulas.length !== parsed.rows || cols !== parsed.columns) return { ok:false, outcome:'dimension-mismatch', error:'formula matrix dimensions must exactly match target range', expected:{rows:parsed.rows,columns:parsed.columns}, actual:{rows:formulas.length,columns:cols} }
  if (parsed.cellCount > maxCells) return { ok:false, outcome:'range-too-large', error:'target range exceeds the configured cell limit', cellCount:parsed.cellCount, maxCells }
  return { ok:true, parsed }
}

function formulaWriterCommand(sheetName, rangeAddress, formulas, maxCells) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra) for(var k in extra)x[k]=extra[k]; return x }
  function parseCell(cell){ var m=String(cell||'').replace(/\$/g,'').match(/^([A-Za-z]+)(\d+)$/); if(!m)return null; var col=0; for(var i=0;i<m[1].length;i++)col=col*26+m[1].toUpperCase().charCodeAt(i)-64; return {row:Number(m[2]),column:col} }
  function parseRange(address){ var p=String(address||'').replace(/\$/g,'').split(':'); if(p.length>2)return null; var a=parseCell(p[0]),b=parseCell(p[1]||p[0]); if(!a||!b||b.row<a.row||b.column<a.column)return null; return {start:a,rows:b.row-a.row+1,columns:b.column-a.column+1} }
  function colLabel(n){ var s=''; while(n>0){n--;s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26)} return s }
  try {
    if(!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=null; try{sheet=Api.GetSheet(sheetName)}catch(_){}
    if(!sheet) return fail('sheet-not-found','the requested worksheet was not found',{sheet:sheetName})
    if(!has(sheet,'GetRange')) return fail('unsupported','ApiWorksheet.GetRange is unavailable',{sheet:sheetName})
    var parsed=parseRange(rangeAddress); if(!parsed)return fail('invalid-range','range must be rectangular A1')
    var count=parsed.rows*parsed.columns; if(count>maxCells)return fail('range-too-large','target range exceeds the configured cell limit',{cellCount:count,maxCells:maxCells})
    if(!Array.isArray(formulas)||formulas.length!==parsed.rows)return fail('dimension-mismatch','formula matrix dimensions must exactly match target range')
    var written=0
    for(var r=0;r<parsed.rows;r++){
      if(!Array.isArray(formulas[r])||formulas[r].length!==parsed.columns)return fail('dimension-mismatch','formula matrix dimensions must exactly match target range')
      for(var c=0;c<parsed.columns;c++){
        var f=formulas[r][c]; if(typeof f!=='string'||f.charAt(0)!=='=')return fail('invalid-formula','every formula cell must begin with =',{row:r+1,column:c+1})
        var address=colLabel(parsed.start.column+c)+(parsed.start.row+r)
        var cell=null; try{cell=sheet.GetRange(address)}catch(_){}
        if(!cell)return fail('range-not-found','a target cell could not be resolved',{cell:address})
        if(!has(cell,'SetFormula'))return fail('unsupported','ApiRange.SetFormula is unavailable',{cell:address})
        try{cell.SetFormula(f);written++}catch(err){return fail('write-error',String(err&&err.message?err.message:err),{cell:address})}
      }
    }
    return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),writtenFormulas:written,cellCount:count}
  } catch(err){ return fail('formula-writer-error',String(err&&err.message?err.message:err),{sheet:sheetName,range:rangeAddress}) }
}

async function writeFormulaInFrame(frame, apiHely, {sheet,range,formulas,maxCells=26000,callbackTimeoutMs=15000}) {
  const valid=validateFormulaMatrix(range,formulas,maxCells)
  if(!valid.ok)return {...valid,source:LIVE_SOURCE,sheet,range}
  const body=`return (${formulaWriterCommand.toString()})(${JSON.stringify(sheet)}, ${JSON.stringify(valid.parsed.address)}, ${JSON.stringify(formulas)}, ${Number(maxCells)});`
  return frame.evaluate(({u,timeout,commandBody})=>new Promise((resolve)=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable on the editor object'})
    let settled=false; const finish=(v)=>{if(!settled){settled=true;resolve(v)}}
    try{editor.callCommand(new Function(commandBody),false,false,(value)=>finish(value===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'}:value))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'formula writer callCommand callback did not arrive in time'}),timeout)
  }),{u:apiHely,timeout:callbackTimeoutMs,commandBody:body})
}

function verifyFormulaMatrix(readBack, formulas) {
  if(!readBack || !readBack.ok || !Array.isArray(readBack.cells)) return {ok:false,outcome:'verification-unavailable',error:'live range read-back did not return a verifiable cell matrix'}
  if(readBack.source !== LIVE_SOURCE) return {ok:false,outcome:'verification-unavailable',error:'verification source is not the live co-edit editor'}
  const mismatches=[]
  for(let r=0;r<formulas.length;r++)for(let c=0;c<formulas[r].length;c++){
    const cell=readBack.cells[r]&&readBack.cells[r][c]
    const actual=cell&&typeof cell.formula==='string'?cell.formula:null
    const status=cell?(cell.formulaStatus || cell.dataType || null):null
    const provenFormula=cell && (cell.formulaStatus ? cell.formulaStatus==='formula' : cell.dataType==='formula')
    if(!provenFormula || actual!==formulas[r][c]) mismatches.push({row:r+1,column:c+1,expected:formulas[r][c],actual,status})
  }
  return mismatches.length?{ok:false,outcome:'verification-mismatch',mismatches}:{ok:true,outcome:'verified'}
}

async function verifyFormulaRangeInFrame(frame, apiHely, {sheet,range,formulas,maxCells=26000,callbackTimeoutMs=15000}) {
  const readBack=await rangeReader.readRangeInFrame(frame,apiHely,{sheet,range,maxCells,callbackTimeoutMs})
  const verification=verifyFormulaMatrix(readBack,formulas)
  return {readBack,verification}
}

async function writeFormulaLive({url,user,pass,fileId,sheet,range,formulas,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}) {
  const valid=validateFormulaMatrix(range,formulas,maxCells); if(!valid.ok)return {...valid,source:LIVE_SOURCE,sheet,range}
  const loaded=loadPlaywright(); if(!loaded.ok)return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium}=loaded.pw; const browser=await chromium.launch()
  try{
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}); const page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]); await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed (browser remained on login page)'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.waitForTimeout(22000)
    const frame=page.frames().find((f)=>/spreadsheeteditor/.test(f.url())); if(!frame)return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>{if((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')return 'window.Asc.editor'; if(window.editor&&typeof window.editor.callCommand==='function')return 'window.editor'; return null})
    if(!apiHely)return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable on known editor objects'}
    const write=await writeFormulaInFrame(frame,apiHely,{sheet,range,formulas,maxCells,callbackTimeoutMs}); if(!write.ok)return {...write,editor:'spreadsheeteditor',apiHely}
    const {readBack,verification}=await verifyFormulaRangeInFrame(frame,apiHely,{sheet,range,formulas,maxCells,callbackTimeoutMs})
    return {...write,verified:verification.ok,verification,readBack,editor:'spreadsheeteditor',apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

module.exports={validateFormulaMatrix,formulaWriterCommand,writeFormulaInFrame,verifyFormulaMatrix,verifyFormulaRangeInFrame,writeFormulaLive}
