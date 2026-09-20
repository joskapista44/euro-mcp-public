'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-editor-page-api-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const probe=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var rx=/(page|print|margin|fit|scale|paper|header|footer|break|worksheet|sheet)/i,seen={},hits=[]
  function scan(label,o){if(!o)return;var chain=[o];try{var p=Object.getPrototypeOf(o);if(p)chain.push(p)}catch(_){}
   for(var ci=0;ci<chain.length;ci++){var names=[];try{names=Object.getOwnPropertyNames(chain[ci])}catch(_){};for(var i=0;i<names.length;i++){var n=names[i];if(n==='constructor'||!rx.test(n))continue;var k=label+'.'+n;if(seen[k])continue;seen[k]=1;var t='unknown';try{t=typeof o[n]}catch(_){t='getter-error'};hits.push({path:k,type:t})}}
  }
  var editor=(typeof Api!=='undefined'&&Api&&Api.editor)?Api.editor:null;scan('Api',Api);scan('Api.editor',editor);if(typeof Asc!=='undefined')scan('Asc',Asc);if(typeof AscCommonExcel!=='undefined')scan('AscCommonExcel',AscCommonExcel)
  hits.sort(function(a,b){return a.path.localeCompare(b.path)});return {ok:true,hasApiEditor:!!editor,hasWb:!!(editor&&editor.wb),hasWbModel:!!(editor&&editor.wbModel),hasWorkbook:!!(editor&&editor.workbook),hits:hits}
 }catch(err){return {ok:false,outcome:'editor-probe-error',error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return probe?.ok?{ok:true,outcome:'xlsx-editor-page-api-probe',authority:'LIVE_VERIFY',noOp:true,probe}:{ok:false,outcome:'xlsx-editor-page-api-probe-failed',authority:'LIVE_READ',noOp:true,probe}
 });console.log('XLSX EDITOR PAGE API PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX EDITOR PAGE API PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
