'use strict'
const assert=require('assert')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const {executeLayoutTaskInPersistentSession}=require('./xlsx-persistent-layout.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-layout-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')},common={fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50}
function d(r){return {outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,applied:r.applied,final:r.final,receipt:r.receipt}}
;(async()=>{const tag=String(Date.now()).slice(-7),sheet=`EURO LYT ${tag}`;const setup=await executeAgentTaskInPersistentSession({...common,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:D4',values:[['Long heading',2,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16]]}]}});assert.equal(setup.ok,true)
const ops=[{range:'A1:A4',type:'column.width',width:18},{range:'A2:D2',type:'row.height',height:24},{range:'B1:B4',type:'columns.hidden',hidden:true},{range:'A3:D3',type:'rows.hidden',hidden:true}]
for(const x of ops){const task={operations:[{intent:'layout_range',sheet,...x}]};const first=await executeLayoutTaskInPersistentSession({...common,task});console.log('LAYOUT FIRST',x.type,JSON.stringify(d(first),null,2));assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.writes,1);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);const retry=await executeLayoutTaskInPersistentSession({...common,task});console.log('LAYOUT RETRY',x.type,JSON.stringify(d(retry),null,2));assert.equal(retry.ok,true);assert.equal(retry.outcome,'xlsx-layout-task-already-satisfied');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)}
console.log('XLSX PERSISTENT LAYOUT LIVE ACCEPTANCE: PASS')})().catch(e=>{console.error(e);process.exitCode=1})
