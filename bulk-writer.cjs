'use strict'

const { parseA1Range } = require('./range-reader.cjs')

function validateMatrix(matrix, name) {
  if (!Array.isArray(matrix) || matrix.length === 0) return { ok: false, error: `${name} must be a non-empty 2D array` }
  if (!Array.isArray(matrix[0]) || matrix[0].length === 0) return { ok: false, error: `${name} must contain non-empty rows` }
  const columns = matrix[0].length
  for (let r = 0; r < matrix.length; r++) if (!Array.isArray(matrix[r]) || matrix[r].length !== columns) return { ok: false, error: `${name} must be rectangular` }
  return { ok: true, rows: matrix.length, columns }
}

function validateWritePayload({ range, values, formulas, maxCells = 26000 }) {
  const parsed = parseA1Range(range)
  if (!parsed) return { ok: false, outcome: 'invalid-range', error: 'range must be a rectangular A1 address such as A1:H50' }
  const vm = validateMatrix(values, 'values')
  if (!vm.ok) return { ok: false, outcome: 'invalid-values', error: vm.error }
  if (vm.rows !== parsed.rows || vm.columns !== parsed.columns) return { ok: false, outcome: 'dimension-mismatch', error: 'values matrix dimensions must exactly match target range', expected: { rows: parsed.rows, columns: parsed.columns }, actual: { rows: vm.rows, columns: vm.columns } }
  if (parsed.cellCount > maxCells) return { ok: false, outcome: 'range-too-large', error: 'target range exceeds the configured cell limit', cellCount: parsed.cellCount, maxCells }
  if (formulas !== undefined && formulas !== null) {
    const fm = validateMatrix(formulas, 'formulas')
    if (!fm.ok) return { ok: false, outcome: 'invalid-formulas', error: fm.error }
    if (fm.rows !== parsed.rows || fm.columns !== parsed.columns) return { ok: false, outcome: 'dimension-mismatch', error: 'formulas matrix dimensions must exactly match target range', expected: { rows: parsed.rows, columns: parsed.columns }, actual: { rows: fm.rows, columns: fm.columns } }
    for (let r = 0; r < fm.rows; r++) for (let c = 0; c < fm.columns; c++) { const f = formulas[r][c]; if (f !== null && f !== undefined && (typeof f !== 'string' || !f.startsWith('='))) return { ok: false, outcome: 'invalid-formula', error: 'non-null formula cells must be strings beginning with =', row: r + 1, column: c + 1 } }
  }
  return { ok: true, parsed }
}

function bulkWriterCommand(sheetName, rangeAddress, values, formulas, maxCells) {
  function has(o,n){return !!o&&typeof o[n]==='function'}
  function fail(outcome,error,extra){var x={ok:false,outcome:outcome,source:'live-coedit-editor',error:error};if(extra)for(var k in extra)x[k]=extra[k];return x}
  function pc(cell){var m=String(cell||'').replace(/\$/g,'').match(/^([A-Za-z]+)(\d+)$/);if(!m)return null;var col=0;for(var i=0;i<m[1].length;i++)col=col*26+m[1].toUpperCase().charCodeAt(i)-64;return{row:Number(m[2]),column:col}}
  function pr(a){var p=String(a||'').replace(/\$/g,'').split(':');if(p.length>2)return null;var x=pc(p[0]),y=pc(p[1]||p[0]);if(!x||!y||y.row<x.row||y.column<x.column)return null;return{start:x,rows:y.row-x.row+1,columns:y.column-x.column+1}}
  function cl(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
  function shape(m){if(!Array.isArray(m)||!m.length||!Array.isArray(m[0])||!m[0].length)return null;var cols=m[0].length;for(var r=0;r<m.length;r++)if(!Array.isArray(m[r])||m[r].length!==cols)return null;return{rows:m.length,columns:cols}}
  try {
    if(!has(Api,'GetSheet'))return fail('unsupported','Api.GetSheet is unavailable')
    var sheet=Api.GetSheet(sheetName);if(!sheet)return fail('sheet-not-found','the requested worksheet was not found',{sheet:sheetName})
    var parsed=pr(rangeAddress),vm=shape(values),fm=formulas==null?null:shape(formulas);if(!parsed||!vm)return fail('invalid-payload','invalid range or values')
    if(vm.rows!==parsed.rows||vm.columns!==parsed.columns||(fm&&(fm.rows!==parsed.rows||fm.columns!==parsed.columns)))return fail('dimension-mismatch','matrix dimensions must exactly match target range')
    var count=parsed.rows*parsed.columns;if(count>maxCells)return fail('range-too-large','target range exceeds configured cell limit',{cellCount:count,maxCells:maxCells})
    var writtenValues=0,writtenFormulas=0,cleared=0
    for(var r=0;r<parsed.rows;r++)for(var c=0;c<parsed.columns;c++){
      var address=cl(parsed.start.column+c)+(parsed.start.row+r),cell=sheet.GetRange(address);if(!cell)return fail('range-not-found','target cell could not be resolved',{cell:address})
      var formula=fm?formulas[r][c]:null
      if(formula!=null){
        if(typeof formula!=='string'||formula.charAt(0)!=='=')return fail('invalid-formula','formula cells must begin with =',{cell:address})
        if(!('Formula' in cell))return fail('unsupported','ApiRange.Formula property is unavailable',{cell:address})
        try{cell.Formula=formula;writtenFormulas++}catch(err){return fail('write-error',String(err&&err.message||err),{cell:address})}
        if(!has(cell,'GetFormula'))return fail('unsupported','ApiRange.GetFormula is unavailable for formula verification',{cell:address})
        var actual=cell.GetFormula();var expectedNorm=String(formula).replace(/^=\s*/,'='),actualNorm=String(actual||'').replace(/^=\s*/,'=')
        if(actualNorm!==expectedNorm)return fail('formula-write-not-verified','formula identity readback mismatch',{cell:address,expected:formula,actual:actual})
        continue
      }
      var value=values[r][c]
      if(value==null){if(has(cell,'Clear')){cell.Clear();cleared++}else if(has(cell,'SetValue')){cell.SetValue('');cleared++}else return fail('unsupported','neither ApiRange.Clear nor ApiRange.SetValue is available',{cell:address})}
      else{if(!has(cell,'SetValue'))return fail('unsupported','ApiRange.SetValue is unavailable',{cell:address});cell.SetValue(value);writtenValues++}
    }
    return{ok:true,outcome:'ok',source:'live-coedit-editor',sheet:sheetName,range:String(rangeAddress).toUpperCase(),rows:parsed.rows,columns:parsed.columns,cellCount:count,writtenValues:writtenValues,writtenFormulas:writtenFormulas,cleared:cleared}
  }catch(err){return fail('bulk-writer-error',String(err&&err.message||err),{sheet:sheetName,range:rangeAddress})}
}

async function writeBulkInFrame(frame, apiHely, { sheet, range, values, formulas = null, maxCells = 26000, callbackTimeoutMs = 15000 }) {
  const valid=validateWritePayload({range,values,formulas,maxCells});if(!valid.ok)return{...valid,source:'live-coedit-editor',sheet,range}
  const body=`return (${bulkWriterCommand.toString()})(${JSON.stringify(sheet)}, ${JSON.stringify(valid.parsed.address)}, ${JSON.stringify(values)}, ${JSON.stringify(formulas)}, ${Number(maxCells)});`
  return frame.evaluate(({u,timeout,commandBody})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand is unavailable on the editor object'});let settled=false;const finish=v=>{if(!settled){settled=true;resolve(v)}};try{editor.callCommand(new Function(commandBody),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor',error:'callCommand callback returned undefined'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message||err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor',error:'bulk writer callCommand callback did not arrive in time'}),timeout)}),{u:apiHely,timeout:callbackTimeoutMs,commandBody:body})
}

async function writeBulkLive({url,user,pass,fileId,sheet,range,values,formulas=null,loadPlaywright,timeoutMs=60000,callbackTimeoutMs=15000,maxCells=26000}){
  const valid=validateWritePayload({range,values,formulas,maxCells});if(!valid.ok)return{...valid,source:'live-coedit-editor',sheet,range};const loaded=loadPlaywright();if(!loaded.ok)return{ok:false,outcome:'nem-mert',source:'live-coedit-editor',error:loaded.indok};const{chromium}=loaded.pw,browser=await chromium.launch();try{const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage();await page.goto(`${url}/login`,{waitUntil:'domcontentloaded',timeout:timeoutMs});await page.fill('#user',user);await page.fill('#password',pass);await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:timeoutMs}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500);if(/\/login/.test(page.url()))return{ok:false,outcome:'auth',source:'live-coedit-editor',error:'login did not succeed'};await page.goto(`${url}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:timeoutMs});await page.waitForTimeout(22000);const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)return{ok:false,outcome:'nem-nyilt-meg',source:'live-coedit-editor',error:'spreadsheeteditor frame did not open'};const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function'?'window.editor':null));if(!apiHely)return{ok:false,outcome:'nincs-api',source:'live-coedit-editor',error:'callCommand unavailable'};return{...(await writeBulkInFrame(frame,apiHely,{sheet,range,values,formulas,maxCells,callbackTimeoutMs})),editor:'spreadsheeteditor',apiHely}}finally{await browser.close().catch(()=>{})}}
module.exports={validateMatrix,validateWritePayload,bulkWriterCommand,writeBulkInFrame,writeBulkLive}
