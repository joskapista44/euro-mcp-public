'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs'),setup=require('./xlsx-persistent-print-setup.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-setup-live-acceptance');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet=`EURO PRINTSET ${String(Date.now()).slice(-7)}`
 const fixture=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);if(!c.ok)return c;const w=await api.writeRangeVerified({sheet,range:'A1:H30',values:Array.from({length:30},(_,r)=>Array.from({length:8},(_,c)=>r===0?'C'+(c+1):(r*10+c))),formulas:Array.from({length:30},()=>Array(8).fill(null))});return w.ok?{ok:true,outcome:'print-setup-fixture-ready',authority:'LIVE_VERIFY',noOp:false}:w});console.log('PRINT SETUP FIXTURE',JSON.stringify(fixture,null,2));assert.equal(fixture.ok,true);assert.equal(fixture.persistentSession?.writes,2);assert.equal(fixture.persistentSession?.persistenceBarrier?.ok,true)
 for(const [label,op,check] of [
  ['FIT',{intent:'set_print_setup',sheet,mode:'fit_to_pages',fitToWidth:1,fitToHeight:1},s=>s.fitToWidth===1&&s.fitToHeight===1],
  ['SCALE',{intent:'set_print_setup',sheet,mode:'scale',scale:85},s=>Math.abs(s.scale-85)<.01],
  ['A4',{intent:'set_print_setup',sheet,mode:'page_size',width:210,height:297},s=>Math.abs(s.width-210)<.01&&Math.abs(s.height-297)<.01]
 ]){
  const task={operations:[op]},first=await setup.executePrintSetupTaskInPersistentSession({...options,task});console.log('PRINT SETUP '+label+' FIRST',JSON.stringify(first,null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(check(first.wholeTaskVerification.state),true)
  const retry=await setup.executePrintSetupTaskInPersistentSession({...options,task});console.log('PRINT SETUP '+label+' RETRY',JSON.stringify(retry,null,2));assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(check(retry.wholeTaskVerification.state),true)
 }
 console.log('XLSX PERSISTENT PRINT SETUP LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
