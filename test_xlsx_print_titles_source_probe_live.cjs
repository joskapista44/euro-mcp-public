'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-source-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var a=Api.GetActiveSheet(),po=a&&a.worksheet&&a.worksheet.PagePrintOptions
  function src(n){try{var f=po&&po[n];if(typeof f!=='function')return null;var s='';try{s=''+f}catch(_){try{s=f.toString()}catch(__){s='[native/uninspectable]'}};return s.length>4000?s.slice(0,4000):s}catch(err){return 'ERROR:'+String(err&&err.message||err)}}
  return {ok:!!po,sheet:a&&a.GetName?a.GetName():null,setHeight:src('asc_setPrintTitlesHeight'),setWidth:src('asc_setPrintTitlesWidth'),init:src('initPrintTitles'),getHeight:src('asc_getPrintTitlesHeight'),getWidth:src('asc_getPrintTitlesWidth')}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-print-titles-source-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-print-titles-source-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PRINT TITLES SOURCE PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PRINT TITLES SOURCE PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
