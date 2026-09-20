'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-options-bridge-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var s=Api.GetActiveSheet(),out={ok:true,sheet:s&&s.GetName?s.GetName():null}
  function desc(label,o){if(!o)return {available:false};var names=[];try{names=Object.getOwnPropertyNames(o).concat(Object.getOwnPropertyNames(Object.getPrototypeOf(o)||{}))}catch(_){};var keep=names.filter(function(n){return /(model|worksheet|sheet|page|print|range|workbook|ws)/i.test(n)}).filter(function(n,i,a){return a.indexOf(n)===i}).sort();return {available:true,type:typeof o,keys:keep.slice(0,120)}}
  out.activeSheet=desc('active',s);out.activeSheetApi=desc('api',s&&s.Api);out.activeSheetWorksheet=desc('ws',s&&s.worksheet);out.activeSheetWs=desc('ws2',s&&s.ws);out.activeSheetModel=desc('model',s&&s.model)
  out.known={worksheet:!!(s&&s.worksheet),ws:!!(s&&s.ws),model:!!(s&&s.model),Api:!!(s&&s.Api)}
  return out
 }catch(err){return {ok:false,outcome:'bridge-probe-error',error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-page-options-bridge-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-page-options-bridge-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PAGE OPTIONS BRIDGE PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PAGE OPTIONS BRIDGE PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
