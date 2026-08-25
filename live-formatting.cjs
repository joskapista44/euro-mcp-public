'use strict'

const LIVE_SOURCE = 'live-coedit-editor'

function validateRgb(v, name) {
  if (v == null) return null
  if (!Array.isArray(v) || v.length !== 3 || v.some((n) => !Number.isInteger(n) || n < 0 || n > 255)) return `${name} must be [r,g,b] integers in 0..255`
  return null
}

function validateFormatSpec(spec) {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return { ok: false, outcome: 'invalid-format', error: 'format must be an object' }
  const allowed = new Set(['bold','italic','fontName','fontSize','fontColor','fillColor','border','alignHorizontal','alignVertical','wrap','numberFormat'])
  for (const k of Object.keys(spec)) if (!allowed.has(k)) return { ok: false, outcome: 'invalid-format', error: `unknown format property: ${k}` }
  for (const k of ['bold','italic','wrap']) if (spec[k] != null && typeof spec[k] !== 'boolean') return { ok: false, outcome: 'invalid-format', error: `${k} must be boolean` }
  if (spec.fontName != null && (typeof spec.fontName !== 'string' || !spec.fontName.trim())) return { ok: false, outcome: 'invalid-format', error: 'fontName must be a non-empty string' }
  if (spec.fontSize != null && (typeof spec.fontSize !== 'number' || spec.fontSize <= 0 || spec.fontSize > 409)) return { ok: false, outcome: 'invalid-format', error: 'fontSize must be > 0 and <= 409' }
  const rgbErr = validateRgb(spec.fontColor, 'fontColor') || validateRgb(spec.fillColor, 'fillColor') || (spec.border && validateRgb(spec.border.color, 'border.color'))
  if (rgbErr) return { ok: false, outcome: 'invalid-format', error: rgbErr }
  if (spec.alignHorizontal != null && !['left','right','center','justify'].includes(spec.alignHorizontal)) return { ok: false, outcome: 'invalid-format', error: 'invalid horizontal alignment' }
  if (spec.alignVertical != null && !['center','bottom','top','distributed','justify'].includes(spec.alignVertical)) return { ok: false, outcome: 'invalid-format', error: 'invalid vertical alignment' }
  if (spec.numberFormat != null && (typeof spec.numberFormat !== 'string' || !spec.numberFormat.length)) return { ok: false, outcome: 'invalid-format', error: 'numberFormat must be a non-empty string' }
  if (spec.border != null) {
    if (typeof spec.border !== 'object' || Array.isArray(spec.border)) return { ok: false, outcome: 'invalid-format', error: 'border must be an object' }
    if (!['Top','Bottom','Left','Right','InsideHorizontal','InsideVertical','DiagonalDown','DiagonalUp'].includes(spec.border.index)) return { ok: false, outcome: 'invalid-format', error: 'invalid border.index' }
    if (typeof spec.border.style !== 'string' || !spec.border.style) return { ok: false, outcome: 'invalid-format', error: 'border.style must be a non-empty string' }
  }
  if (!Object.keys(spec).length) return { ok: false, outcome: 'invalid-format', error: 'format must contain at least one property' }
  return { ok: true }
}

function formattingCommand(sheetName, rangeAddress, spec) {
  function has(o, n) { return !!o && typeof o[n] === 'function' }
  function fail(outcome, error, extra) { var x = { ok:false, outcome:outcome, source:'live-coedit-editor', error:error }; if (extra) for (var k in extra) x[k]=extra[k]; return x }
  function color(rgb) { return rgb == null ? null : Api.CreateColorFromRGB(rgb[0], rgb[1], rgb[2]) }
  try {
    if (!has(Api, 'GetSheet')) return fail('unsupported', 'Api.GetSheet is unavailable')
    var sheet = null; try { sheet = Api.GetSheet(sheetName) } catch (_) {}
    if (!sheet || !has(sheet, 'GetRange')) return fail('sheet-not-found', 'worksheet/range API unavailable', { sheet:sheetName })
    var range = null; try { range = sheet.GetRange(rangeAddress) } catch (_) {}
    if (!range) return fail('range-not-found', 'target range could not be resolved')
    var applied = []
    function apply(key, method, value) {
      if (!has(range, method)) throw { unsupported:true, message:'ApiRange.' + method + ' is unavailable', capability:key }
      var r = range[method](value); if (r === false) throw { message:'ApiRange.' + method + ' returned false', capability:key }
      applied.push(key)
    }
    if (spec.bold !== undefined) apply('bold','SetBold',spec.bold)
    if (spec.italic !== undefined) apply('italic','SetItalic',spec.italic)
    if (spec.fontName !== undefined) apply('fontName','SetFontName',spec.fontName)
    if (spec.fontSize !== undefined) apply('fontSize','SetFontSize',spec.fontSize)
    if (spec.fontColor !== undefined) apply('fontColor','SetFontColor',color(spec.fontColor))
    if (spec.fillColor !== undefined) apply('fillColor','SetFillColor',color(spec.fillColor))
    if (spec.alignHorizontal !== undefined) apply('alignHorizontal','SetAlignHorizontal',spec.alignHorizontal)
    if (spec.alignVertical !== undefined) apply('alignVertical','SetAlignVertical',spec.alignVertical)
    if (spec.wrap !== undefined) apply('wrap','SetWrap',spec.wrap)
    if (spec.numberFormat !== undefined) apply('numberFormat','SetNumberFormat',spec.numberFormat)
    if (spec.border !== undefined) {
      if (!has(range,'SetBorders')) return fail('unsupported','ApiRange.SetBorders is unavailable',{ capability:'border' })
      var br = range.SetBorders(spec.border.index, spec.border.style, color(spec.border.color)); if (br === false) return fail('format-error','ApiRange.SetBorders returned false',{ capability:'border' })
      applied.push('border')
    }
    var verified = {}
    if (spec.numberFormat !== undefined && has(range,'GetNumberFormat')) verified.numberFormat = range.GetNumberFormat()
    if (spec.wrap !== undefined && has(range,'GetWrapText')) verified.wrap = range.GetWrapText()
    return { ok:true, outcome:'ok', source:'live-coedit-editor', sheet:sheetName, range:String(rangeAddress).toUpperCase(), applied:applied, verified:verified }
  } catch (err) {
    if (err && err.unsupported) return fail('unsupported',err.message,{ capability:err.capability })
    return fail('format-error',String(err && err.message ? err.message : err),{ capability:err && err.capability ? err.capability : null })
  }
}

function layoutCommand(sheetName, rangeAddress, options) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error,extra){ var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error}; if(extra)for(var k in extra)x[k]=extra[k]; return x }
  try {
    if (!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName); if(!sheet || !has(sheet,'GetRange')) return fail('sheet-not-found','worksheet/range API unavailable')
    var range=sheet.GetRange(rangeAddress); if(!range) return fail('range-not-found','target range could not be resolved')
    var applied=[]
    if(options.columnWidth!==undefined){ if(!has(range,'SetColumnWidth')) return fail('unsupported','ApiRange.SetColumnWidth is unavailable',{capability:'columnWidth'}); range.SetColumnWidth(options.columnWidth); applied.push('columnWidth') }
    if(options.rowHeight!==undefined){ if(!has(range,'SetRowHeight')) return fail('unsupported','ApiRange.SetRowHeight is unavailable',{capability:'rowHeight'}); var rh=range.SetRowHeight(options.rowHeight); if(rh===false)return fail('layout-error','SetRowHeight returned false'); applied.push('rowHeight') }
    if(options.autofitRows || options.autofitColumns){ if(!has(range,'AutoFit')) return fail('unsupported','ApiRange.AutoFit is unavailable',{capability:'autofit'}); range.AutoFit(!!options.autofitRows,!!options.autofitColumns); applied.push('autofit') }
    return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),applied:applied}
  } catch(err){ return fail('layout-error',String(err && err.message ? err.message : err)) }
}

function structureCommand(sheetName, rangeAddress, operation, across) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error){ return {ok:false,outcome:outcome,source:'live-coedit-editor',error:error} }
  try {
    if(!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName); if(!sheet || !has(sheet,'GetRange')) return fail('sheet-not-found','worksheet/range API unavailable')
    var range=sheet.GetRange(rangeAddress); if(!range) return fail('range-not-found','target range could not be resolved')
    var method=operation==='merge'?'Merge':'UnMerge'; if(!has(range,method)) return fail('unsupported','ApiRange.'+method+' is unavailable')
    var result=operation==='merge'?range.Merge(!!across):range.UnMerge(); if(result===false) return fail('structure-error','ApiRange.'+method+' returned false')
    return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),operation:operation,across:operation==='merge'?!!across:null}
  } catch(err){ return fail('structure-error',String(err && err.message ? err.message : err)) }
}

function freezeCommand(sheetName, action, value) {
  function has(o,n){ return !!o && typeof o[n] === 'function' }
  function fail(outcome,error){ return {ok:false,outcome:outcome,source:'live-coedit-editor',error:error} }
  try {
    if(!has(Api,'GetSheet')) return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName); if(!sheet || !has(sheet,'GetFreezePanes')) return fail('unsupported','ApiWorksheet.GetFreezePanes is unavailable')
    var fp=sheet.GetFreezePanes(); if(!fp) return fail('unsupported','freeze panes object is unavailable')
    if(action==='unfreeze'){ if(!has(fp,'Unfreeze'))return fail('unsupported','ApiFreezePanes.Unfreeze is unavailable'); fp.Unfreeze() }
    else if(action==='rows'){ if(!has(fp,'FreezeRows'))return fail('unsupported','ApiFreezePanes.FreezeRows is unavailable'); fp.FreezeRows(value) }
    else if(action==='columns'){ if(!has(fp,'FreezeColumns'))return fail('unsupported','ApiFreezePanes.FreezeColumns is unavailable'); fp.FreezeColumns(value) }
    else if(action==='at'){ if(!has(fp,'FreezeAt') || !has(sheet,'GetRange'))return fail('unsupported','ApiFreezePanes.FreezeAt is unavailable'); var r=sheet.GetRange(value); if(!r)return fail('range-not-found','freeze target range could not be resolved'); fp.FreezeAt(r) }
    else return fail('invalid-freeze-action','action must be rows, columns, at or unfreeze')
    var location=null; if(has(fp,'GetLocation')){ var loc=fp.GetLocation(); if(loc && has(loc,'GetAddress')) location=loc.GetAddress() }
    return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:sheetName,action:action,value:value===undefined?null:value,location:location}
  } catch(err){ return fail('freeze-error',String(err && err.message ? err.message : err)) }
}

async function runInLiveEditor({ url,user,pass,fileId,loadPlaywright,command,args,timeoutMs=60000,callbackTimeoutMs=15000 }) {
  const loaded=loadPlaywright(); if(!loaded.ok)return {ok:false,outcome:'nem-mert',source:LIVE_SOURCE,error:loaded.indok}
  const { chromium }=loaded.pw; const browser=await chromium.launch()
  try {
    const ctx=await browser.newContext({viewport:{width:1400,height:900}}); const page=await ctx.newPage()
    await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.fill('#user',user); await page.fill('#password',pass)
    await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]); await page.waitForTimeout(2500)
    if(/\/login/.test(page.url()))return {ok:false,outcome:'auth',source:LIVE_SOURCE,error:'login did not succeed'}
    await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs}); await page.waitForTimeout(22000)
    const frame=page.frames().find((f)=>/spreadsheeteditor/.test(f.url())); if(!frame)return {ok:false,outcome:'nem-nyilt-meg',source:LIVE_SOURCE,error:'spreadsheeteditor frame did not open'}
    const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null)
    if(!apiHely)return {ok:false,outcome:'nincs-api',source:LIVE_SOURCE,error:'callCommand is unavailable'}
    const body=`return (${command.toString()}).apply(null, ${JSON.stringify(args)});`
    const result=await frame.evaluate(({u,commandBody,timeout})=>new Promise((resolve)=>{ const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor; let done=false; const finish=(v)=>{if(!done){done=true;resolve(v)}}; try{ editor.callCommand(new Function(commandBody),false,false,(v)=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'}:v)) }catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})} setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'formatting callback timed out'}),timeout) }),{u:apiHely,commandBody:body,timeout:callbackTimeoutMs})
    return {...result,editor:'spreadsheeteditor',apiHely}
  } finally { await browser.close().catch(()=>{}) }
}

function formatRangeLive(common){ const v=validateFormatSpec(common.format); if(!v.ok)return {...v,source:LIVE_SOURCE}; return runInLiveEditor({...common,command:formattingCommand,args:[common.sheet,common.range,common.format]}) }
function layoutRangeLive(common){ return runInLiveEditor({...common,command:layoutCommand,args:[common.sheet,common.range,{columnWidth:common.columnWidth,rowHeight:common.rowHeight,autofitRows:common.autofitRows,autofitColumns:common.autofitColumns}]}) }
function mergeRangeLive(common){ return runInLiveEditor({...common,command:structureCommand,args:[common.sheet,common.range,common.operation,common.across]}) }
function freezePanesLive(common){ return runInLiveEditor({...common,command:freezeCommand,args:[common.sheet,common.action,common.value]}) }

module.exports={ LIVE_SOURCE,validateFormatSpec,formattingCommand,layoutCommand,structureCommand,freezeCommand,runInLiveEditor,formatRangeLive,layoutRangeLive,mergeRangeLive,freezePanesLive }
