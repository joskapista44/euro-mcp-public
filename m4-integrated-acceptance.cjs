'use strict'

const { runOperationInFrame } = require('./workbook-ops.cjs')
const { writeBulkInFrame } = require('./bulk-writer.cjs')
const { readRangeInFrame } = require('./range-reader.cjs')
const { formattingCommand } = require('./live-formatting.cjs')
const { runStructureInFrame } = require('./live-structure.cjs')
const { writeCrossSheetInFrame } = require('./cross-sheet-formulas.cjs')
const { runConditionalFormattingInFrame } = require('./live-conditional-formatting.cjs')

function status(result) {
  if (!result) return 'FAIL'
  if (result.verification && result.verification.status) return result.verification.status
  if (result.verification && result.verification.outcome) return String(result.verification.outcome).toUpperCase()
  return result.ok ? 'PASS' : 'FAIL'
}
function firstCell(read) { return read && read.ok && read.cells && read.cells[0] ? read.cells[0][0] : null }
function loadPlaywright() { for (const p of [process.env.EURO_PLAYWRIGHT_PATH,'playwright','/home/user/marveen/node_modules/playwright'].filter(Boolean)) { try { return require(p) } catch (_) {} } throw new Error('Playwright is unavailable') }
async function callFormatting(frame, apiHely, sheet, range, spec, timeout=15000) {
  const body=`return (${formattingCommand.toString()})(${JSON.stringify(sheet)},${JSON.stringify(range)},${JSON.stringify(spec)});`
  return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const e=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};try{e.callCommand(new Function(body),false,v=>finish(v===undefined?{ok:false,outcome:'empty-callback',source:'live-coedit-editor'}:v))}catch(err){finish({ok:false,outcome:'callcommand-error',source:'live-coedit-editor',error:String(err&&err.message?err.message:err)})}setTimeout(()=>finish({ok:false,outcome:'callback-timeout',source:'live-coedit-editor'}),timeout)}),{u:apiHely,body,timeout})
}

async function runIntegrated(frame, apiHely) {
  const suffix=String(Date.now()).slice(-8), data=`M4DATA_${suffix}`, summary=`M4SUM_${suffix}`
  const steps=[],cleanup=[]; let dataCreated=false,summaryCreated=false
  const push=(milestone,name,result,forced)=>{const s=forced||status(result);steps.push({milestone,name,status:s,result});return s}
  try {
    let r=await runOperationInFrame(frame,apiHely,{type:'sheet.create',name:data});push('M4.4','create-data-sheet',r);dataCreated=!!r.ok;if(!r.ok)throw new Error('create data sheet failed')
    r=await runOperationInFrame(frame,apiHely,{type:'sheet.create',name:summary});push('M4.4','create-summary-sheet',r);summaryCreated=!!r.ok;if(!r.ok)throw new Error('create summary sheet failed')

    const values=[['Customer','Principal','Rate','Status'],['Alpha',100000,0.12,'Active'],['Beta',150000,0.1,'Active'],['Gamma',80000,0.15,'Watch']]
    r=await writeBulkInFrame(frame,apiHely,{sheet:data,range:'A1:D4',values});push('core','seed-business-data',r);if(!r.ok)throw new Error('seed failed')
    const rd=await readRangeInFrame(frame,apiHely,{sheet:data,range:'A1:D4',maxCells:32});const seedPass=rd&&rd.ok&&rd.cells&&rd.cells.length===4;push('core','seed-live-readback',rd,seedPass?'PASS':'FAIL');if(!seedPass)throw new Error('seed readback failed')

    r=await callFormatting(frame,apiHely,data,'A1:D1',{bold:true,fillColor:[220,230,241],alignHorizontal:'center'});push('M4.1','header-format',r);if(status(r)==='FAIL')throw new Error('formatting failed')
    r=await callFormatting(frame,apiHely,data,'B2:B4',{numberFormat:'#,##0'});push('M4.2','principal-number-format',r);if(status(r)==='FAIL')throw new Error('number format failed')
    r=await callFormatting(frame,apiHely,data,'C2:C4',{numberFormat:'0.00%'});push('M4.2','rate-number-format',r);if(status(r)==='FAIL')throw new Error('percent format failed')

    // M4.3 layout: use the already proven direct live AutoFit API and require the command to execute.
    const layoutBody=`return (function(s,r){var sh=Api.GetSheet(s),x=sh&&sh.GetRange(r);if(!x||typeof x.AutoFit!=='function')return {ok:false,outcome:'unsupported',source:'live-coedit-editor'};x.AutoFit(false,true);return {ok:true,outcome:'ok',source:'live-coedit-editor',sheet:s,range:r};})(${JSON.stringify(data)},'A:D');`
    r=await frame.evaluate(({u,b})=>new Promise(resolve=>{const e=u==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(b),false,v=>resolve(v))}),{u:apiHely,b:layoutBody});push('M4.3','autofit-columns',r);if(!r||!r.ok)throw new Error('autofit failed')

    const formula=`='${data}'!B2*(1+'${data}'!C2)`
    r=await writeCrossSheetInFrame(frame,apiHely,{sheet:summary,range:'A1',formulas:[[formula]],maxCells:4});push('M4.4','cross-sheet-formula',r,r&&r.ok&&r.verified?'PASS':'FAIL');if(!r||!r.ok||!r.verified)throw new Error('formula write failed')
    let fr=await readRangeInFrame(frame,apiHely,{sheet:summary,range:'A1',maxCells:4});let cell=firstCell(fr),expected=112000;let formulaPass=!!cell&&!/#REF!/i.test(cell.formula||'')&&(Number(cell.rawValue)===expected||Number(cell.displayText)===expected);push('M4.4','formula-live-value',fr,formulaPass?'PASS':'FAIL');if(!formulaPass)throw new Error('formula verification failed')

    r=await runStructureInFrame(frame,apiHely,{type:'range.merge',sheet:summary,range:'C1:D1'});push('M4.4','merge-summary-title-area',r);if(status(r)!=='PASS')throw new Error('merge failed')
    r=await runStructureInFrame(frame,apiHely,{type:'range.unmerge',sheet:summary,range:'C1:D1'});push('M4.4','unmerge-summary-title-area',r);if(status(r)!=='PASS')throw new Error('unmerge failed')

    steps.push({milestone:'M4.5',name:'excel-table',status:'DEFERRED',reason:'EuroOffice 9.3.4-hotfix.1 lacks machine-verifiable ListObject API; recorded in M45-DEFERRED.md'})

    r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.inspect',sheet:data,range:'B2:B4'});push('M4.6','cf-initial-inspect',r);if(!r.ok||r.count!==0)throw new Error('conditional-formatting range not isolated')
    r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.add',sheet:data,range:'B2:B4',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'120000',fillColor:[255,255,0],priority:1}});push('M4.6','cf-add-high-principal',r);if(status(r)!=='PASS')throw new Error('conditional formatting add failed')
    r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.inspect',sheet:data,range:'B2:B4'});push('M4.6','cf-live-readback',r,r&&r.ok&&r.count===1?'PASS':'FAIL');if(!r.ok||r.count!==1)throw new Error('conditional formatting readback failed')
    r=await runConditionalFormattingInFrame(frame,apiHely,{type:'cf.delete',sheet:data,range:'B2:B4',index:0});push('M4.6','cf-cleanup-rule',r);if(status(r)!=='PASS')throw new Error('conditional formatting cleanup failed')
  } catch(err) { steps.push({milestone:'M4','name':'acceptance-exception',status:'FAIL',error:String(err&&err.message?err.message:err)}) }
  finally {
    if(summaryCreated){const r=await runOperationInFrame(frame,apiHely,{type:'sheet.delete',sheet:summary});cleanup.push({sheet:summary,status:r&&r.ok?'PASS':'FAIL',result:r})}
    if(dataCreated){const r=await runOperationInFrame(frame,apiHely,{type:'sheet.delete',sheet:data});cleanup.push({sheet:data,status:r&&r.ok?'PASS':'FAIL',result:r})}
  }
  const required=steps.filter(s=>s.status!=='DEFERRED'), testOutcome=required.some(s=>s.status==='FAIL')?'FAIL':required.some(s=>s.status==='UNKNOWN')?'UNKNOWN':'PASS'
  const cleanupOutcome=cleanup.length===2&&cleanup.every(x=>x.status==='PASS')?'PASS':'FAIL'
  return {milestone:'M4 Integrated Acceptance',source:'live-coedit-editor',outcome:testOutcome==='PASS'&&cleanupOutcome==='PASS'?'PASS':testOutcome==='FAIL'||cleanupOutcome==='FAIL'?'FAIL':'UNKNOWN',testOutcome,cleanupOutcome,humanObservationRequired:false,deferred:[{milestone:'M4.5',reason:'runtime capability; recheck on EuroOffice based on ONLYOFFICE 9.4+'}],sheets:{data,summary},steps,cleanup}
}

async function main(){const base=process.env.EURO_NC_BASE_URL,fileId=process.env.EURO_NC_FILE_ID,user=process.env.EURO_NC_USER,password=process.env.EURO_NC_PASSWORD;if(!base||!fileId||!user||!password)throw new Error('EURO_NC_BASE_URL, EURO_NC_FILE_ID, EURO_NC_USER and EURO_NC_PASSWORD are required');const {chromium}=loadPlaywright(),browser=await chromium.launch();try{const ctx=await browser.newContext({viewport:{width:1400,height:900}}),page=await ctx.newPage();await page.goto(`${base}/login`,{waitUntil:'domcontentloaded',timeout:60000});await page.fill('#user',user);await page.fill('#password',password);await Promise.all([page.waitForNavigation({waitUntil:'domcontentloaded',timeout:60000}).catch(()=>null),page.click('button[type=submit], input[type=submit]')]);await page.waitForTimeout(2500);if(/\/login/.test(page.url()))throw new Error('Nextcloud login failed');await page.goto(`${base}/index.php/apps/eurooffice/${fileId}`,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForTimeout(22000);const frame=page.frames().find(f=>/spreadsheeteditor/.test(f.url()));if(!frame)throw new Error('spreadsheeteditor frame not found');const apiHely=await frame.evaluate(()=>((window.Asc||{}).editor&&typeof window.Asc.editor.callCommand==='function')?'window.Asc.editor':(window.editor&&typeof window.editor.callCommand==='function')?'window.editor':null);if(!apiHely)throw new Error('callCommand unavailable');console.log(JSON.stringify({...await runIntegrated(frame,apiHely),editor:'spreadsheeteditor',apiHely},null,2))}finally{await browser.close()}}

module.exports={status,firstCell,runIntegrated}
if(require.main===module)main().catch(err=>{console.error(JSON.stringify({milestone:'M4 Integrated Acceptance',outcome:'launcher-error',error:String(err&&err.message?err.message:err)},null,2));process.exitCode=1})
