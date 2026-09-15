'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')

function env(name){const v=process.env[name];if(!v)throw new Error(`${name} is required`);return v}
const fileId=env('EURO_XLSX_FILE_ID')
const credentials={url:env('EURO_NC_URL'),user:env('EURO_NC_USER'),pass:env('EURO_NC_PASS')}
const sheet=process.env.EURO_XLSX_TEST_SHEET||'EURO Persistent LIVE'
const stamp=process.env.EURO_XLSX_TEST_VALUE||'persistent-live-v1'
const task={operations:[
 {intent:'create_sheet',name:sheet},
 {intent:'write_range',sheet,range:'A1:B2',values:[[stamp,17],['formula',null]],formulas:[[null,null],[null,'=B1*2']]}
]}
async function run(label){const t0=Date.now();const r=await executeAgentTaskInPersistentSession({fileId,task,credentials,timeoutMs:30000,pollMs:50});console.log(`=== ${label} ===`);console.log(JSON.stringify(r,null,2));console.log(`WALL_MS=${Date.now()-t0}`);assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.wholeTaskVerification?.ok,true);assert.equal(r.persistentSession?.oneEditorSession,true);return r}
;(async()=>{const first=await run('FIRST TASK');assert.ok(first.receipt.some(x=>x.status==='APPLIED')||first.noOp===true);if(!first.noOp){assert.ok(first.persistentSession.persistenceBarrier?.ok);assert.ok(first.persistentSession.writes>=1)}const second=await run('SECOND IDENTICAL TASK');assert.equal(second.noOp,true);assert.ok(second.receipt.every(x=>x.status==='NO_OP_SATISFIED'));assert.equal(second.persistentSession.writes,0);assert.equal(second.persistentSession.persistenceBarrier,null);console.log('XLSX PERSISTENT LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e&&e.stack||e);process.exitCode=1})
