'use strict'
const assert=require('assert')
const {executeCopyChainTaskInPersistentSession}=require('./xlsx-persistent-copy-chain.cjs')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-copy-chain-retry-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const common={fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50}
;(async()=>{
 const tag=String(Date.now()).slice(-7),source=`EURO CCR A ${tag}`,copy=`EURO CCR B ${tag}`,final=`EURO CCR C ${tag}`
 const setup=await executeAgentTaskInPersistentSession({...common,task:{operations:[{intent:'create_sheet',name:source},{intent:'write_range',sheet:source,range:'A1:B2',values:[[11,22],[33,44]],formulas:[[null,null],[null,'=A1+B1']]}]}});assert.equal(setup.ok,true);assert.equal(setup.authority,'LIVE_VERIFY')
 const task={operations:[{intent:'copy_sheet',sheet:source,name:copy},{intent:'write_range',sheet:copy,range:'A1',values:[[99]]},{intent:'rename_sheet',sheet:copy,name:final}]}
 const first=await executeCopyChainTaskInPersistentSession({...common,task});console.log('FIRST',JSON.stringify({outcome:first.outcome,authority:first.authority,noOp:first.noOp,writes:first.persistentSession?.writes,barrier:first.persistentSession?.persistenceBarrier?.ok,receipt:first.receipt},null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await executeCopyChainTaskInPersistentSession({...common,task});console.log('RETRY',JSON.stringify({outcome:retry.outcome,authority:retry.authority,noOp:retry.noOp,writes:retry.persistentSession?.writes,barrier:retry.persistentSession?.persistenceBarrier,classification:retry.retryClassification,receipt:retry.receipt},null,2));assert.equal(retry.ok,true);assert.equal(retry.outcome,'xlsx-copy-chain-task-already-satisfied');assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.ok(retry.receipt.every(x=>x.status==='NO_OP_SATISFIED'))
 console.log('XLSX PERSISTENT COPY CHAIN RETRY LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
