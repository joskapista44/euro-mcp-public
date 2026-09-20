'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-advanced-api-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var a=Api.GetActiveSheet(),ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions,ps=po&&po.asc_getPageSetup?po.asc_getPageSetup():null,hf=ps&&ps.asc_getHeaderFooter?ps.asc_getHeaderFooter():null
  function methods(o,rx){if(!o)return [];var names=[];try{names=Object.getOwnPropertyNames(o).concat(Object.getOwnPropertyNames(Object.getPrototypeOf(o)||{}))}catch(_){};return names.filter(function(n){return rx.test(n)}).filter(function(n,i,x){return x.indexOf(n)===i}).sort()}
  function call(o,n){try{return o&&typeof o[n]==='function'?{available:true,value:o[n]()}:{available:false}}catch(err){return {available:true,error:String(err&&err.message||err)}}}
  return {ok:!!ws,sheet:a&&a.GetName?a.GetName():null,
   worksheetMethods:methods(ws,/(print|title|break|page|area)/i),
   optionMethods:methods(po,/(center|grid|heading|margin|setup|print)/i),
   setupMethods:methods(ps,/(fit|scale|width|height|orientation|header|footer|page|paper|black|draft|error|comment|order)/i),
   headerFooterMethods:methods(hf,/(header|footer|first|odd|even|margin|scale|align)/i),
   values:{horizontalCentered:call(po,'asc_getHorizontalCentered'),verticalCentered:call(po,'asc_getVerticalCentered'),printTitlesHeight:call(po,'asc_getPrintTitlesHeight'),printTitlesWidth:call(po,'asc_getPrintTitlesWidth'),headerFooter:call(ps,'asc_getHeaderFooter')}
  }
 }catch(err){return {ok:false,outcome:'advanced-print-probe-error',error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-print-advanced-api-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-print-advanced-api-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PRINT ADVANCED API PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PRINT ADVANCED API PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
