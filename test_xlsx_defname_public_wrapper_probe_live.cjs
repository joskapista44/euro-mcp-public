'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-defname-public-wrapper-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{const p=await api.session.frame.evaluate(({where})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;const body=`try{
 var out={},fn=Api.AddDefName;out.addDefNameSource=String(fn);out.own=[];try{out.own=Object.getOwnPropertyNames(fn)}catch(_){}
 out.fnProps={};for(var i=0;i<out.own.length;i++){var k=out.own[i];try{var v=fn[k];out.fnProps[k]={type:typeof v,value:(typeof v==='string'||typeof v==='number'||typeof v==='boolean')?v:null,source:typeof v==='function'?String(v):null}}catch(_){}}
 var apiOwn=[];try{apiOwn=Object.getOwnPropertyNames(Api)}catch(_){};out.apiPrivateCandidates=apiOwn.filter(function(x){return /(def|name|private)/i.test(x)}).sort();return {ok:true,out:out}
 }catch(err){return {ok:false,error:String(err&&err.stack||err)}}`;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere});return p&&p.ok?{ok:true,outcome:'xlsx-defname-public-wrapper-probe',authority:'LIVE_VERIFY',noOp:true,probe:p.out}:{ok:false,outcome:'xlsx-defname-public-wrapper-probe-failed',authority:'LIVE_READ',probe:p}})
 console.log('XLSX DEFNAME PUBLIC WRAPPER PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);console.log('XLSX DEFNAME PUBLIC WRAPPER PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
