'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs'),titles=require('./xlsx-persistent-print-titles.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
async function task(spec){return persistent.withPersistentXlsxSession(options,async api=>{const pre=await titles.runCommand(api.session,spec,false);if(pre.ok&&pre.noOp)return {ok:true,outcome:'print-titles-already-satisfied',authority:'LIVE_VERIFY',noOp:true,state:pre.state};const a=await titles.runCommand(api.session,spec,true);if(!a.ok)return a;api.session.markWrite();return {ok:true,outcome:'print-titles-live-verified',authority:'LIVE_VERIFY',noOp:false,state:a.state}})}
;(async()=>{
 const sheet=`EURO TITLES ${String(Date.now()).slice(-7)}`
 const fixture=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);return c.ok?{ok:true,outcome:'print-titles-fixture-ready',authority:'LIVE_VERIFY',noOp:false}:c});assert.equal(fixture.ok,true);assert.equal(fixture.persistentSession?.persistenceBarrier?.ok,true)
 const rows={sheet,axis:'rows',from:1,to:2};let r=await task(rows);console.log('PRINT TITLES ROWS FIRST',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.state.rows,'$1:$2');assert.equal(r.persistentSession?.persistenceBarrier?.ok,true);r=await task(rows);console.log('PRINT TITLES ROWS RETRY',JSON.stringify(r,null,2));assert.equal(r.noOp,true);assert.equal(r.persistentSession?.writes,0)
 const cols={sheet,axis:'columns',from:1,to:2};r=await task(cols);console.log('PRINT TITLES COLS FIRST',JSON.stringify(r,null,2));assert.equal(r.ok,true);assert.equal(r.state.columns,'$A:$B');assert.equal(r.state.rows,'$1:$2');assert.equal(r.persistentSession?.persistenceBarrier?.ok,true);r=await task(cols);console.log('PRINT TITLES COLS RETRY',JSON.stringify(r,null,2));assert.equal(r.noOp,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.state.columns,'$A:$B');assert.equal(r.state.rows,'$1:$2')
 console.log('XLSX PRINT TITLES LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
