'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-title-editor-surface-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{const r=await persistent.withPersistentXlsxSession(options,async api=>{
 const p=await api.session.frame.evaluate(()=>{function scan(o,label){if(!o)return {label,available:false};let n=[];try{let x=o,depth=0;while(x&&depth++<4){n=n.concat(Object.getOwnPropertyNames(x));x=Object.getPrototypeOf(x)}}catch(_){};n=[...new Set(n)].filter(x=>/(print|page|def|name|title)/i.test(x)).sort();let src={};for(const k of n){try{src[k]=typeof o[k]==='function'?String(o[k]).slice(0,2500):{type:typeof o[k]}}catch(e){src[k]={error:String(e)}}}return {label,available:true,methods:n,sources:src}}
 const candidates=[['window.editor',window.editor],['window.Asc.editor',window.Asc&&window.Asc.editor],['window.AscCommonExcel',window.AscCommonExcel],['window.Asc',window.Asc]]
 return {ok:true,candidates:candidates.map(([l,o])=>scan(o,l))}
 })
 return p&&p.ok?{ok:true,outcome:'xlsx-print-title-editor-surface-probe',authority:'LIVE_VERIFY',noOp:true,probe:p}:{ok:false,outcome:'xlsx-print-title-editor-surface-probe-failed',authority:'LIVE_READ',probe:p}
 });console.log('XLSX PRINT TITLE EDITOR SURFACE PROBE',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.persistentSession?.writes,0);console.log('XLSX PRINT TITLE EDITOR SURFACE PROBE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
