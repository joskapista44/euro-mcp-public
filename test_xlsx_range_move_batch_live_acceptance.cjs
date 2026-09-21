'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-range-move-batch-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function view(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier,checks:r.wholeTaskVerification?.checks,steps:r.steps?.map(s=>({index:s.index,intent:s.intent,outcome:s.result?.outcome,noOp:s.result?.noOp}))}}
;(async()=>{const tag=String(Date.now()).slice(-7),src='EURO RM SRC '+tag,dst='EURO RM DST '+tag,task={operations:[
{intent:'create_sheet',name:src},{intent:'create_sheet',name:dst},
{intent:'write_range',sheet:src,range:'A1:B3',values:[['Item','Amount'],['Alpha',10],['Beta',20]],formulas:[[null,null],[null,null],[null,null]]},
{intent:'move_range',sheet:src,range:'A1:B3',targetSheet:dst,targetRange:'D4:E6'}
],readbacks:[{sheet:src,range:'A1:B3'},{sheet:dst,range:'D4:E6'}]}
const first=await batch.executeBatchTaskInPersistentSession({...options,task});console.log('RANGE MOVE BATCH FIRST',JSON.stringify(view(first),null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(first.wholeTaskVerification?.checks?.every(x=>x.ok),true)
const retry=await batch.executeBatchTaskInPersistentSession({...options,task});console.log('RANGE MOVE BATCH RETRY',JSON.stringify(view(retry),null,2));assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.wholeTaskVerification?.checks?.every(x=>x.ok),true)
console.log('XLSX RANGE MOVE BATCH LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
