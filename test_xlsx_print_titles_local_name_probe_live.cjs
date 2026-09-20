'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-local-name-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var sheets=Api.GetSheets(),out=[];for(var i=0;i<sheets.length;i++){var a=sheets[i],ws=a&&a.worksheet;if(!ws)continue;var name=a.GetName?a.GetName():null,id=ws.getId?ws.getId():null,wb=ws.workbook,d=null;try{d=wb&&wb.getDefinesNames?wb.getDefinesNames("Print_Titles",id):null}catch(err){d={error:String(err&&err.message||err)}};out.push({sheet:name,id:id,local:d?{name:d.name||null,ref:d.ref||null,localSheetId:d.localSheetId===undefined?null:d.localSheetId}:null})}
  return {ok:true,sheets:out}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-print-titles-local-name-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-print-titles-local-name-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX PRINT TITLES LOCAL NAME PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);console.log('XLSX PRINT TITLES LOCAL NAME PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
