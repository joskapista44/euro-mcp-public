'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
const PLAYWRIGHT=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
process.env.EURO_PLAYWRIGHT_PATH=PLAYWRIGHT
function secret(id){const path='/home/user/marveen/dist/web/vault.js';const v=require(path);const r=v.getSecret(id,'xlsx-persistent-delete-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
async function run(task){return executeAgentTaskInPersistentSession({fileId:FILE_ID,task,credentials,timeoutMs:30000,pollMs:50})}
;(async()=>{const suffix=String(Date.now()).slice(-7),sheet=`EURO Del ${suffix}`;console.log(`SHEET=${sheet}`)
const setup=await run({operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1',values:[['delete-persist']]}]});console.log('=== DELETE SETUP TASK ===');console.log(JSON.stringify(setup,null,2));assert.equal(setup.ok,true);assert.equal(setup.authority,'LIVE_VERIFY');assert.equal(setup.noOp,false);assert.ok(setup.persistentSession?.persistenceBarrier?.ok)
const first=await run({operations:[{intent:'delete_sheet',sheet}]});console.log('=== DELETE APPLIED TASK ===');console.log(JSON.stringify(first,null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.receipt?.[0]?.status,'APPLIED');assert.equal(first.wholeTaskVerification?.checks?.[0]?.status,'PASS');assert.equal(first.persistentSession?.writes,1);assert.ok(first.persistentSession?.persistenceBarrier?.ok)
const second=await run({operations:[{intent:'delete_sheet',sheet}]});console.log('=== DELETE IDEMPOTENCY TASK ===');console.log(JSON.stringify(second,null,2));assert.equal(second.ok,true);assert.equal(second.authority,'LIVE_VERIFY');assert.equal(second.noOp,true);assert.equal(second.receipt?.[0]?.status,'NO_OP_SATISFIED');assert.equal(second.persistentSession?.writes,0);assert.equal(second.persistentSession?.persistenceBarrier,null)
console.log('XLSX PERSISTENT DELETE LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
