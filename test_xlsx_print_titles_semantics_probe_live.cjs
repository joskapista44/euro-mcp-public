'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-semantics-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var a=Api.GetActiveSheet(),ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions
  function own(o){if(!o)return [];var n=[];try{n=Object.getOwnPropertyNames(o).concat(Object.getOwnPropertyNames(Object.getPrototypeOf(o)||{}))}catch(_){};return n.filter(function(x){return /(printTitles|initPrintTitles)/i.test(x)}).filter(function(x,i,z){return z.indexOf(x)===i}).sort()}
  return {ok:!!po,sheet:a&&a.GetName?a.GetName():null,methods:own(po),height:po&&po.asc_getPrintTitlesHeight?po.asc_getPrintTitlesHeight():null,width:po&&po.asc_getPrintTitlesWidth?po.asc_getPrintTitlesWidth():null,
   setterArity:{height:po&&po.asc_setPrintTitlesHeight?po.asc_setPrintTitlesHeight.length:null,width:po&&po.asc_setPrintTitlesWidth?po.asc_setPrintTitlesWidth.length:null,init:po&&po.initPrintTitles?po.initPrintTitles.length:null}}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-print-titles-semantics-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-print-titles-semantics-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PRINT TITLES SEMANTICS PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PRINT TITLES SEMANTICS PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
