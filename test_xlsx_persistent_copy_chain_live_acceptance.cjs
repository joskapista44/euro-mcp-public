'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-copy-chain-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
async function run(task){return executeAgentTaskInPersistentSession({fileId:FILE_ID,task,credentials,timeoutMs:30000,pollMs:50})}
;(async()=>{
 const tag=String(Date.now()).slice(-7),source=`EURO Chain A ${tag}`,copy=`EURO Chain B ${tag}`,final=`EURO Chain C ${tag}`
 console.log(`SOURCE=${source}`);console.log(`COPY=${copy}`);console.log(`FINAL=${final}`)
 const setup={operations:[{intent:'create_sheet',name:source},{intent:'write_range',sheet:source,range:'A1:B2',values:[[11,22],[33,44]],formulas:[[null,null],[null,'=A1+B1']]}]}
 const setupResult=await run(setup);console.log('=== COPY CHAIN SETUP ===');console.log(JSON.stringify(setupResult,null,2));assert.equal(setupResult.ok,true);assert.equal(setupResult.authority,'LIVE_VERIFY')
 const chain={operations:[{intent:'copy_sheet',sheet:source,name:copy},{intent:'write_range',sheet:copy,range:'A1',values:[[99]]},{intent:'rename_sheet',sheet:copy,name:final}]}
 const result=await run(chain);console.log('=== COPY WRITE RENAME CHAIN ===');console.log(JSON.stringify(result,null,2));assert.equal(result.ok,true);assert.equal(result.authority,'LIVE_VERIFY');assert.equal(result.noOp,false);assert.deepEqual(result.receipt.map(x=>x.status),['APPLIED','APPLIED','APPLIED']);assert.equal(result.wholeTaskVerification?.ok,true);assert.equal(result.wholeTaskVerification?.checks?.every(x=>x.status==='PASS'),true);assert.equal(result.persistentSession?.oneEditorSession,true);assert.equal(result.persistentSession?.writes,3);assert.equal(result.persistentSession?.persistenceBarrier?.ok,true)
 console.log('XLSX PERSISTENT COPY CHAIN LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e);process.exitCode=1})
