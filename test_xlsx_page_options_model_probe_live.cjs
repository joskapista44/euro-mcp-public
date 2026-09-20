'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-options-model-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var wb=(typeof Asc!=='undefined'&&Asc.editor&&Asc.editor.wbModel)?Asc.editor.wbModel:null;if(!wb)return {ok:false,outcome:'wb-model-unavailable'}
  var idx=typeof Asc.editor.asc_getActiveWorksheetIndex==='function'?Asc.editor.asc_getActiveWorksheetIndex():0,ws=wb.getWorksheet(idx);if(!ws)return {ok:false,outcome:'model-sheet-unavailable',idx:idx}
  var po=ws.PagePrintOptions,ps=po&&po.asc_getPageSetup?po.asc_getPageSetup():null,pm=po&&po.asc_getPageMargins?po.asc_getPageMargins():null
  function call(o,n){try{return o&&typeof o[n]==='function'?{available:true,value:o[n]()}:{available:false}}catch(err){return {available:true,error:String(err&&err.message||err)}}}
  return {ok:true,index:idx,name:ws.getName?ws.getName():null,pageOptions:!!po,pageSetup:!!ps,pageMargins:!!pm,read:{
   width:call(ps,'asc_getWidth'),height:call(ps,'asc_getHeight'),orientation:call(ps,'asc_getOrientation'),fitToWidth:call(ps,'asc_getFitToWidth'),fitToHeight:call(ps,'asc_getFitToHeight'),scale:call(ps,'asc_getScale'),firstPageNumber:call(ps,'asc_getFirstPageNumber'),
   left:call(pm,'asc_getLeft'),right:call(pm,'asc_getRight'),top:call(pm,'asc_getTop'),bottom:call(pm,'asc_getBottom'),
   gridLines:call(po,'asc_getGridLines'),headings:call(po,'asc_getHeadings'),horizontalCentered:call(po,'asc_getHorizontalCentered'),verticalCentered:call(po,'asc_getVerticalCentered')
  }}
 }catch(err){return {ok:false,outcome:'model-probe-error',error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-page-options-model-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-page-options-model-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PAGE OPTIONS MODEL PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PAGE OPTIONS MODEL PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
