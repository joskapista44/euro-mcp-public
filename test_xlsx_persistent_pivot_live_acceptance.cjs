'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const pivotAgent=require('./xlsx-agent-pivot-task.cjs')
const pivotPersistent=require('./xlsx-persistent-pivot.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-pivot-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,pivotObserved:async(spec,apply)=>{const r=await pivotPersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function createTask(sheet,name){return {operations:[{intent:'create_pivot',name,sourceSheet:sheet,sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:10},{items:['East','B'],expected:30},{items:['West','B'],expected:20}]}]}}
function deleteTask(name,pivotSheet){return {operations:[{intent:'delete_pivot_sheet',name,pivotSheet}]}}
async function firstInvocation(api,sheet,name){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'pivot-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const values=[['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]
 const written=await api.writeRangeVerified({sheet,range:'A1:C5',values});if(!written.ok)return {ok:false,outcome:'pivot-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 return pivotAgent.executePivotTask({task:createTask(sheet,name),api:liveApi(api)})
}
function exactCreate(r,name){const s=r?.wholeTaskVerification?.state;return s?.present===true&&s?.name===name&&s?.source==='A1:C5'&&s?.rowFields===1&&s?.columnFields===1&&s?.dataFields===1&&s?.styleName==='PivotStyleMedium2'&&s?.assertions?.every(x=>x.ok&&x.value===x.expected)}
function exactDelete(r,name){const s=r?.wholeTaskVerification?.state;return s?.pivot?.present===false&&s?.pivot?.name===name&&s?.pivotSheetPresent===false}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
;(async()=>{
 const suffix=String(Date.now()).slice(-7),sheet=`EURO PIV ${suffix}`,name=`EURO_P_${suffix}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet,name));console.log('PIVOT CREATE TASK',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(exactCreate(first,name),true);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const pivotSheet=first.wholeTaskVerification.state.parentSheet;assert.equal(typeof pivotSheet,'string')
 const retry=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:createTask(sheet,name),api:liveApi(api)}));console.log('PIVOT CREATE RETRY',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(exactCreate(retry,name),true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 const deleted=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:deleteTask(name,pivotSheet),api:liveApi(api)}));console.log('PIVOT DELETE TASK',JSON.stringify(d(deleted),null,2))
 assert.equal(deleted.ok,true);assert.equal(deleted.authority,'LIVE_VERIFY');assert.equal(deleted.noOp,false);assert.equal(exactDelete(deleted,name),true);assert.equal(deleted.persistentSession?.writes,1);assert.equal(deleted.persistentSession?.persistenceBarrier?.ok,true)
 const deleteRetry=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:deleteTask(name,pivotSheet),api:liveApi(api)}));console.log('PIVOT DELETE RETRY',JSON.stringify(d(deleteRetry),null,2))
 assert.equal(deleteRetry.ok,true);assert.equal(deleteRetry.authority,'LIVE_VERIFY');assert.equal(deleteRetry.noOp,true);assert.equal(exactDelete(deleteRetry,name),true);assert.equal(deleteRetry.persistentSession?.writes,0);assert.equal(deleteRetry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT PIVOT LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
