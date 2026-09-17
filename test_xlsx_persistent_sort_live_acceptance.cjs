'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const sortAgent=require('./xlsx-agent-sort-task.cjs')
const sortPersistent=require('./xlsx-persistent-sort.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-sort-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,sortObserved:async(spec,apply)=>{const r=await sortPersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function task(sheet){return {operations:[{intent:'sort_range',sheet,range:'A1:C5',keyRange:'A1:A5',order:'asc',hasHeaders:true}]}}
async function firstInvocation(api,sheet){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'sort-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const values=[['Name','Score','Group'],['Delta',40,'B'],['Alpha',10,'A'],['Charlie',30,'A'],['Bravo',20,'B']]
 const written=await api.writeRangeVerified({sheet,range:'A1:C5',values});if(!written.ok)return {ok:false,outcome:'sort-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 return sortAgent.executeSortTask({task:task(sheet),api:liveApi(api)})
}
async function retryInvocation(api,sheet){return sortAgent.executeSortTask({task:task(sheet),api:liveApi(api)})}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification}}
;(async()=>{
 const sheet=`EURO SORT ${String(Date.now()).slice(-7)}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet))
 console.log('SORT TASK 1 (ONE EDITOR SESSION)',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.deepStrictEqual(first.wholeTaskVerification?.keys,['Alpha','Bravo','Charlie','Delta']);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,api=>retryInvocation(api,sheet))
 console.log('SORT TASK 2 IDEMPOTENT REINVOCATION',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.deepStrictEqual(retry.wholeTaskVerification?.keys,['Alpha','Bravo','Charlie','Delta']);assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT SORT LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
