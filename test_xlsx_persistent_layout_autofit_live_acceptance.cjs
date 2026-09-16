'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const {executeLayoutTaskInPersistentSession}=require('./xlsx-persistent-layout.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-layout-autofit-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')},common={fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,applied:r.applied,final:r.final,wholeTaskVerification:r.wholeTaskVerification,receipt:r.receipt}}
;(async()=>{const tag=String(Date.now()).slice(-7),sheet=`EURO AF ${tag}`;const setup=await executeAgentTaskInPersistentSession({...common,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:C3',values:[['A very long heading that should force a wider column','x','y'],['another long value for autofit verification','short','short'],['third long value','z','z']]}]}});assert.equal(setup.ok,true)
const ops=[{range:'A1:A3',type:'columns.autofit'},{range:'A1:C3',type:'rows.autofit'}]
for(const x of ops){const task={operations:[{intent:'layout_range',sheet,...x}]};const first=await executeLayoutTaskInPersistentSession({...common,task});console.log('AUTOFIT FIRST',x.type,JSON.stringify(d(first),null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.writes,1);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(Number.isFinite(first.wholeTaskVerification?.actualDimension),true)}
console.log('XLSX PERSISTENT LAYOUT AUTOFIT LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
