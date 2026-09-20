'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-deep-api-inventory');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const inventory=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
  var s=Api.GetActiveSheet(),hits=[],seen=new Set(),queue=[{path:'sheet',obj:s,depth:0}]
  var rx=/(print|page|margin|header|footer|break|fit|scale|paper|orient|setup|title|width|height)/i
  while(queue.length){var q=queue.shift(),o=q.obj;if(!o||q.depth>3)continue;var names=[];try{names=Object.getOwnPropertyNames(o)}catch(_){};try{var p=Object.getPrototypeOf(o);if(p)names=names.concat(Object.getOwnPropertyNames(p))}catch(_){}
   for(var i=0;i<names.length;i++){var n=names[i];if(n==='constructor')continue;var key=q.path+'.'+n;if(seen.has(key))continue;seen.add(key);var t='unknown',v;try{v=o[n];t=typeof v}catch(err){t='getter-error'}
    if(rx.test(n))hits.push({path:key,type:t})
    if(q.depth<3&&v&&typeof v==='object'&&/(print|page|view|sheet|model|workbook|ws|setup)/i.test(n))queue.push({path:key,obj:v,depth:q.depth+1})
   }
  }
  hits.sort(function(a,b){return a.path.localeCompare(b.path)});return {ok:true,sheet:typeof s.GetName==='function'?s.GetName():null,hits:hits}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere})
 return inventory?.ok?{ok:true,outcome:'xlsx-print-deep-api-inventory',authority:'LIVE_VERIFY',noOp:true,inventory}:{ok:false,outcome:'xlsx-print-deep-api-inventory-failed',authority:'LIVE_READ',noOp:true,inventory}
 });console.log('XLSX PRINT DEEP API INVENTORY',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null);console.log('XLSX PRINT DEEP API INVENTORY: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
