'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs'),centering=require('./xlsx-persistent-print-centering.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-centering-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet=`EURO CENTER ${String(Date.now()).slice(-7)}`
 const fixture=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);return c.ok?{ok:true,outcome:'centering-fixture-ready',authority:'LIVE_VERIFY',noOp:false}:c});assert.equal(fixture.ok,true);assert.equal(fixture.persistentSession?.persistenceBarrier?.ok,true)
 const spec={sheet,horizontal:true,vertical:true}
 const first=await persistent.withPersistentXlsxSession(options,async api=>{const pre=await centering.runCommand(api.session,spec,false);if(pre.ok&&pre.noOp)return {ok:true,outcome:'centering-already-satisfied',authority:'LIVE_VERIFY',noOp:true,state:pre.state};const a=await centering.runCommand(api.session,spec,true);if(!a.ok)return a;api.session.markWrite();return {ok:true,outcome:'centering-live-verified',authority:'LIVE_VERIFY',noOp:false,state:a.state}})
 console.log('PRINT CENTERING FIRST',JSON.stringify(first,null,2));assert.equal(first.ok,true);assert.equal(first.state.horizontal,true);assert.equal(first.state.vertical,true);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,async api=>{const pre=await centering.runCommand(api.session,spec,false);return pre.ok&&pre.noOp?{ok:true,outcome:'centering-already-satisfied',authority:'LIVE_VERIFY',noOp:true,state:pre.state}:pre})
 console.log('PRINT CENTERING RETRY',JSON.stringify(retry,null,2));assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.state.horizontal,true);assert.equal(retry.state.vertical,true)
 console.log('XLSX PRINT CENTERING LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
