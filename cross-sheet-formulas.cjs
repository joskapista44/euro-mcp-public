'use strict'

const formulaWriter = require('./formula-writer.cjs')

const LIVE_SOURCE = 'live-coedit-editor'

function quoteSheetName(name) {
  if (typeof name !== 'string' || !name.length) return null
  return `'${name.replace(/'/g, "''")}'`
}

function qualifyA1(sheet, ref) {
  if (typeof ref !== 'string' || !/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(ref)) return {ok:false,outcome:'invalid-reference',error:'reference must be an A1 cell or range'}
  const q = quoteSheetName(sheet)
  if (!q) return {ok:false,outcome:'invalid-sheet',error:'sheet must be a non-empty string'}
  return {ok:true,formulaRef:`${q}!${ref}`}
}

function scanSheetReferences(formula) {
  if (typeof formula !== 'string' || formula[0] !== '=') return {ok:false,outcome:'invalid-formula',error:'formula must begin with ='}
  const refs=[]
  let i=1, inString=false
  while(i<formula.length){
    const ch=formula[i]
    if(ch==='"'){
      if(inString && formula[i+1]==='"'){i+=2;continue}
      inString=!inString;i++;continue
    }
    if(inString){i++;continue}
    if(ch==="'"){
      let j=i+1,name=''
      while(j<formula.length){
        if(formula[j]==="'"){
          if(formula[j+1]==="'"){name+="'";j+=2;continue}
          break
        }
        name+=formula[j++]
      }
      if(j>=formula.length || formula[j+1]!=='!') return {ok:false,outcome:'invalid-sheet-reference',error:'quoted sheet name is not correctly terminated before !'}
      refs.push({sheet:name,quoted:true,start:i,end:j+2}); i=j+2; continue
    }
    const m=formula.slice(i).match(/^([A-Za-z_][A-Za-z0-9_.]*)!/)
    if(m){refs.push({sheet:m[1],quoted:false,start:i,end:i+m[0].length});i+=m[0].length;continue}
    i++
  }
  if(inString)return {ok:false,outcome:'invalid-formula',error:'unterminated string literal'}
  return {ok:true,references:refs,sheets:[...new Set(refs.map(r=>r.sheet))]}
}

function validateCrossSheetMatrix(formulas) {
  if(!Array.isArray(formulas)||!formulas.length||!Array.isArray(formulas[0])||!formulas[0].length)return {ok:false,outcome:'invalid-formulas',error:'formulas must be a non-empty 2D array'}
  const cols=formulas[0].length, sheets=new Set(), details=[]
  for(let r=0;r<formulas.length;r++){
    if(!Array.isArray(formulas[r])||formulas[r].length!==cols)return {ok:false,outcome:'invalid-formulas',error:'formulas must be rectangular'}
    for(let c=0;c<cols;c++){
      const s=scanSheetReferences(formulas[r][c]); if(!s.ok)return {...s,row:r+1,column:c+1}
      if(!s.references.length)return {ok:false,outcome:'missing-cross-sheet-reference',error:'every M3.4 formula must contain at least one sheet-qualified reference',row:r+1,column:c+1}
      for(const x of s.sheets)sheets.add(x)
      details.push({row:r+1,column:c+1,references:s.references})
    }
  }
  return {ok:true,sheets:[...sheets],details}
}

async function preflightReferencedSheets(frame,apiHely,sheets,callbackTimeoutMs=15000){
  return frame.evaluate(({u,names,timeout})=>new Promise(resolve=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable on the editor object'})
    let settled=false; const finish=v=>{if(!settled){settled=true;resolve(v)}}
    const body=`return (function(names){function has(o,n){return !!o&&typeof o[n]==='function'};if(!has(Api,'GetSheet'))return {ok:false,outcome:'unsupported',source:'live-coedit-editor',error:'Api.GetSheet is unavailable'};var missing=[];for(var i=0;i<names.length;i++){var sh=null;try{sh=Api.GetSheet(names[i])}catch(_){};if(!sh)missing.push(names[i])}return missing.length?{ok:false,outcome:'referenced-sheet-not-found',source:'live-coedit-editor',missingSheets:missing}:{ok:true,outcome:'ok',source:'live-coedit-editor',validatedSheets:names};})(${JSON.stringify(names)});`
    try{editor.callCommand(new Function(body),false,false,value=>finish(value===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'sheet preflight callback returned undefined'}:value))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'sheet preflight callback did not arrive in time'}),timeout)
  }),{u:apiHely,names:sheets,timeout:callbackTimeoutMs})
}

async function writeCrossSheetInFrame(frame,apiHely,{sheet,range,formulas,maxCells=26000,callbackTimeoutMs=15000}){
  const cross=validateCrossSheetMatrix(formulas); if(!cross.ok)return {...cross,source:LIVE_SOURCE,sheet,range}
  const writerValid=formulaWriter.validateFormulaMatrix(range,formulas,maxCells); if(!writerValid.ok)return {...writerValid,source:LIVE_SOURCE,sheet,range}
  const preflight=await preflightReferencedSheets(frame,apiHely,cross.sheets,callbackTimeoutMs); if(!preflight.ok)return {...preflight,targetSheet:sheet,range}
  const write=await formulaWriter.writeFormulaInFrame(frame,apiHely,{sheet,range,formulas,maxCells,callbackTimeoutMs}); if(!write.ok)return {...write,referencedSheets:cross.sheets}
  const {readBack,verification}=await formulaWriter.verifyFormulaRangeInFrame(frame,apiHely,{sheet,range,formulas,maxCells,callbackTimeoutMs})
  return {...write,referencedSheets:cross.sheets,validatedSheets:preflight.validatedSheets,verified:verification.ok,verification,readBack}
}

async function writeCrossSheetLive({url,user,pass,fileId,sheet,range,formulas,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}){
  const cross=validateCrossSheetMatrix(formulas); if(!cross.ok)return {...cross,source:LIVE_SOURCE,sheet,range}
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
    const result=await writeCrossSheetInFrame(frame,apiHely,{sheet,range,formulas,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally {await browser.close().catch(()=>{})}
}

module.exports={quoteSheetName,qualifyA1,scanSheetReferences,validateCrossSheetMatrix,preflightReferencedSheets,writeCrossSheetInFrame,writeCrossSheetLive}
