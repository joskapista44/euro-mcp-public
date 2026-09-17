'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const filterAgent=require('./xlsx-agent-filter-task.cjs')
const filterPersistent=require('./xlsx-persistent-filter.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-filter-clear-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,filterObserved:async(spec,apply)=>{const r=await filterPersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function setTask(sheet){return {operations:[{intent:'filter_range',sheet,range:'A1:C5',field:3,criteria1:'A',operator:'xlOr'}]}}
function clearTask(sheet){return {operations:[{intent:'clear_filter',sheet,range:'A1:C5'}]}}
async function firstInvocation(api,sheet){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'filter-clear-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const values=[['Name','Score','Group'],['Delta',40,'B'],['Alpha',10,'A'],['Charlie',30,'A'],['Bravo',20,'B']]
 const written=await api.writeRangeVerified({sheet,range:'A1:C5',values});if(!written.ok)return {ok:false,outcome:'filter-clear-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 const prepared=await filterAgent.executeFilterTask({task:setTask(sheet),api:liveApi(api)});if(!prepared.ok)return {ok:false,outcome:'filter-clear-fixture-filter-failed',authority:prepared.authority||'LIVE_READ',prepared}
 return filterAgent.executeFilterTask({task:clearTask(sheet),api:liveApi(api)})
}
function exactClear(r){const s=r?.wholeTaskVerification?.state;return Array.isArray(s?.filters)&&s.filters.length===0&&(!s.present||s.range==='A1:C5')}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
;(async()=>{
 const sheet=`EURO FCLR ${String(Date.now()).slice(-7)}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet))
 console.log('FILTER CLEAR TASK 1 (ONE EDITOR SESSION)',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(exactClear(first),true);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,4);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,api=>filterAgent.executeFilterTask({task:clearTask(sheet),api:liveApi(api)}))
 console.log('FILTER CLEAR TASK 2 IDEMPOTENT REINVOCATION',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(exactClear(retry),true);assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT FILTER CLEAR LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
