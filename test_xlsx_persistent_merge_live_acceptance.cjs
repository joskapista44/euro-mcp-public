'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const {executeMergeTaskInPersistentSession}=require('./xlsx-persistent-merge.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-merge-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')},common={fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,receipt:r.receipt}}
;(async()=>{const tag=String(Date.now()).slice(-7),sheet=`EURO MRG ${tag}`;const setup=await executeAgentTaskInPersistentSession({...common,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:B2',values:[['MERGE ME',''],['','']]}]}});assert.equal(setup.ok,true)
for(const intent of ['merge_range','unmerge_range']){const task={operations:[{intent,sheet,range:'A1:B2'}]};const first=await executeMergeTaskInPersistentSession({...common,task});console.log('MERGE FIRST',intent,JSON.stringify(d(first),null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.writes,1);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(first.wholeTaskVerification?.topLeftPreserved,true);const retry=await executeMergeTaskInPersistentSession({...common,task});console.log('MERGE RETRY',intent,JSON.stringify(d(retry),null,2));assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)}
console.log('XLSX PERSISTENT MERGE LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
