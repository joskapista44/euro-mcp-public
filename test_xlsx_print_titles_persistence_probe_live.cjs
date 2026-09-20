'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-persistence-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var sheets=Api.GetSheets(),a=sheets[sheets.length-1],ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions,wb=ws&&ws.workbook
  function def(){try{var d=wb&&wb.getDefinesNames?wb.getDefinesNames("Print_Titles",ws.getId()):null;return d?{name:d.name||null,ref:d.ref||null,sheetId:ws.getId()}:null}catch(err){return {error:String(err&&err.message||err)}}}
  var before={sheet:a&&a.GetName?a.GetName():null,height:po&&po.asc_getPrintTitlesHeight?po.asc_getPrintTitlesHeight():null,width:po&&po.asc_getPrintTitlesWidth?po.asc_getPrintTitlesWidth():null,definedName:def()}
  if(po&&po.initPrintTitles)po.initPrintTitles()
  var after={height:po&&po.asc_getPrintTitlesHeight?po.asc_getPrintTitlesHeight():null,width:po&&po.asc_getPrintTitlesWidth?po.asc_getPrintTitlesWidth():null,definedName:def()}
  return {ok:!!ws,before:before,afterInit:after}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-print-titles-persistence-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-print-titles-persistence-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PRINT TITLES PERSISTENCE PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PRINT TITLES PERSISTENCE PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
