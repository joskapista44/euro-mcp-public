'use strict'

const { formattingCommand, validateFormatSpec, LIVE_SOURCE } = require('./live-formatting.cjs')
const { readRangeInFrame } = require('./range-reader.cjs')

function digits(n){ return n > 0 ? '.' + '0'.repeat(n) : '' }
function buildNumberFormat(spec={}) {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return {ok:false,outcome:'invalid-number-format',error:'format must be an object'}
  if (typeof spec.mask === 'string' && spec.mask.length) return {ok:true,kind:'custom',mask:spec.mask}
  const kind=spec.kind
  const decimals=spec.decimals === undefined ? 2 : spec.decimals
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 12) return {ok:false,outcome:'invalid-number-format',error:'decimals must be an integer in 0..12'}
  const grouped=spec.grouped === undefined ? true : !!spec.grouped
  const number=(grouped?'#,##0':'0')+digits(decimals)
  if(kind==='decimal') return {ok:true,kind,mask:number}
  if(kind==='percent') return {ok:true,kind,mask:'0'+digits(decimals)+'%'}
  if(kind==='currency') {
    const symbol=typeof spec.symbol==='string' && spec.symbol.length ? spec.symbol : '€'
    return {ok:true,kind,mask:symbol+' '+number}
  }
  if(kind==='date') return {ok:true,kind,mask:spec.pattern || 'yyyy-mm-dd'}
  if(kind==='time') return {ok:true,kind,mask:spec.pattern || 'hh:mm:ss'}
  if(kind==='datetime') return {ok:true,kind,mask:spec.pattern || 'yyyy-mm-dd hh:mm:ss'}
  return {ok:false,outcome:'invalid-number-format',error:'kind must be currency, percent, date, time, datetime, decimal, or provide mask'}
}

function same(a,b){
  if (Object.is(a,b)) return true
  if (Number.isNaN(a) && Number.isNaN(b)) return true
  return false
}

function flatCells(read){
  return read && Array.isArray(read.cells) ? read.cells.flat() : []
}

function verifyNumberFormatTransition(before, write, after, expectedMask) {
  if (!write || write.ok !== true) return {ok:false,outcome:write && write.outcome || 'write-failed',source:LIVE_SOURCE,write,verification:{outcome:'fail',reason:'format write did not succeed'}}
  if (!before || before.ok !== true || !after || after.ok !== true) {
    return {ok:false,outcome:'verification-unknown',source:LIVE_SOURCE,write,before,after,verification:{outcome:'unknown',reason:'live range readback unavailable before or after formatting'}}
  }
  const a=flatCells(before), b=flatCells(after)
  if (a.length !== b.length || a.length===0) return {ok:false,outcome:'verification-unknown',source:LIVE_SOURCE,write,before,after,verification:{outcome:'unknown',reason:'live readback cell shape changed or was empty'}}
  const cells=[]; const mismatches=[]; const unknown=[]
  for(let i=0;i<a.length;i++){
    const pre=a[i], post=b[i]
    if(!pre || !post || pre.address!==post.address){ unknown.push(post&&post.address || pre&&pre.address || String(i)); continue }
    const formatKnown=post.numberFormat !== null && post.numberFormat !== undefined
    const formatPass=formatKnown && post.numberFormat===expectedMask
    const rawPass=same(pre.rawValue,post.rawValue)
    const status=(!formatKnown)?'unknown':(formatPass&&rawPass?'pass':'fail')
    if(status==='fail')mismatches.push(post.address)
    if(status==='unknown')unknown.push(post.address)
    cells.push({address:post.address,status,beforeRawValue:pre.rawValue,rawValue:post.rawValue,value:post.value,displayText:post.displayText,numberFormat:post.numberFormat,checks:{numberFormat:{status:formatKnown?(formatPass?'pass':'fail'):'unknown',expected:expectedMask,actual:post.numberFormat},underlyingValuePreserved:{status:rawPass?'pass':'fail',before:pre.rawValue,after:post.rawValue}}})
  }
  const outcome=mismatches.length?'fail':unknown.length?'unknown':'pass'
  return {ok:outcome==='pass',outcome:outcome==='pass'?'ok':outcome==='fail'?'verification-mismatch':'verification-unknown',source:LIVE_SOURCE,write,verification:{outcome,expectedNumberFormat:expectedMask,cells,mismatches,unknown}}
}

async function callCommandInFrame(frame, apiHely, command, args, callbackTimeoutMs=15000) {
  const body=`return (${command.toString()}).apply(null, ${JSON.stringify(args)});`
  return frame.evaluate(({u,commandBody,timeout})=>new Promise((resolve)=>{
    const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
    if(!editor || typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable'})
    let done=false; const finish=(v)=>{if(!done){done=true;resolve(v)}}
    try{ editor.callCommand(new Function(commandBody),false,false,(v)=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'}:v)) }
    catch(err){ finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)}) }
    setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'number-format callCommand callback timed out'}),timeout)
  }),{u:apiHely,commandBody:body,timeout:callbackTimeoutMs})
}

async function formatNumberInFrame(frame, apiHely, {sheet,range,numberFormat,maxCells=26000,callbackTimeoutMs=15000}){
  const built=typeof numberFormat==='string'?{ok:true,mask:numberFormat,kind:'custom'}:buildNumberFormat(numberFormat)
  if(!built.ok)return {...built,source:LIVE_SOURCE}
  const valid=validateFormatSpec({numberFormat:built.mask}); if(!valid.ok)return {...valid,source:LIVE_SOURCE}
  const before=await readRangeInFrame(frame,apiHely,{sheet,range,maxCells,callbackTimeoutMs})
  if(!before.ok)return {ok:false,outcome:'verification-unknown',source:LIVE_SOURCE,before,verification:{outcome:'unknown',reason:'pre-format live readback unavailable; write was not attempted'}}
  const write=await callCommandInFrame(frame,apiHely,formattingCommand,[sheet,range,{numberFormat:built.mask}],callbackTimeoutMs)
  if(!write.ok)return verifyNumberFormatTransition(before,write,null,built.mask)
  const after=await readRangeInFrame(frame,apiHely,{sheet,range,maxCells,callbackTimeoutMs})
  const result=verifyNumberFormatTransition(before,write,after,built.mask)
  return {...result,kind:built.kind,mask:built.mask,before,after}
}

async function formatNumberRangeLive({url,user,pass,fileId,sheet,range,numberFormat,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}){
  const loaded=loadPlaywright(); if(!loaded.ok)return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const {chromium}=loaded.pw; const browser=await chromium.launch()
  try{
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}); const page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]); await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.waitForTimeout(22000)
    const frame=page.frames().find((f)=>/spreadsheeteditor/.test(f.url())); if(!frame)return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null)
    if(!apiHely)return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable'}
    const result=await formatNumberInFrame(frame,apiHely,{sheet,range,numberFormat,maxCells,callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

module.exports={buildNumberFormat,verifyNumberFormatTransition,callCommandInFrame,formatNumberInFrame,formatNumberRangeLive}
