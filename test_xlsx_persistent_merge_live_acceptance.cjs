'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const mergeAgent=require('./xlsx-agent-merge-task.cjs')
const mergePersistent=require('./xlsx-persistent-merge.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-merge-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function mergeApi(api){return {...api,mergeObserved:async(spec,apply)=>{const r=await mergePersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function task(intent,sheet){return {operations:[{intent,sheet,range:'A1:B2'}]}}
async function firstInvocation(api,sheet){
 const created=await api.createSheetVerified(sheet)
 if(!created.ok)return {ok:false,outcome:'unmerge-fixture-create-failed',authority:created.authority||'LIVE_READ',noOp:false,created}
 const written=await api.writeRangeVerified({sheet,range:'A1:B2',values:[['MERGE ME',''],['','']]})
 if(!written.ok)return {ok:false,outcome:'unmerge-fixture-write-failed',authority:written.authority||'LIVE_READ',noOp:false,written}
 const liveApi=mergeApi(api)
 const merged=await mergeAgent.executeMergeTask({task:task('merge_range',sheet),api:liveApi})
 if(!merged.ok)return {ok:false,outcome:'unmerge-fixture-merge-failed',authority:merged.authority||'LIVE_READ',noOp:false,merged}
 const unmerged=await mergeAgent.executeMergeTask({task:task('unmerge_range',sheet),api:liveApi})
 if(!unmerged.ok)return {ok:false,outcome:'unmerge-task-failed',authority:unmerged.authority||'LIVE_READ',noOp:false,merged,unmerged}
 const final=await mergeAgent.executeMergeTask({task:task('unmerge_range',sheet),api:liveApi})
 if(!final.ok||final.authority!=='LIVE_VERIFY'||final.noOp!==true)return {ok:false,outcome:'unmerge-whole-task-live-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',noOp:false,merged,unmerged,final}
 return {ok:true,outcome:'unmerge-task-live-verified',authority:'LIVE_VERIFY',noOp:false,merged,unmerged,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',outcome:'unmerge-whole-task-live-verified',final}}
}
async function retryInvocation(api,sheet){
 const liveApi=mergeApi(api)
 const result=await mergeAgent.executeMergeTask({task:task('unmerge_range',sheet),api:liveApi})
 if(!result.ok||result.authority!=='LIVE_VERIFY'||result.noOp!==true)return {ok:false,outcome:'unmerge-retry-not-satisfied',authority:result.authority||'LIVE_READ',noOp:false,result}
 const final=await mergeAgent.executeMergeTask({task:task('unmerge_range',sheet),api:liveApi})
 if(!final.ok||final.authority!=='LIVE_VERIFY'||final.noOp!==true)return {ok:false,outcome:'unmerge-retry-whole-task-live-verify-failed',authority:'LIVE_READ',noOp:false,result,final}
 return {ok:true,outcome:'unmerge-task-already-satisfied',authority:'LIVE_VERIFY',noOp:true,result,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',outcome:'unmerge-whole-task-live-verified',final}}
}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification&&{ok:r.wholeTaskVerification.ok,outcome:r.wholeTaskVerification.outcome}}}
;(async()=>{
 const sheet=`EURO UNMRG ${String(Date.now()).slice(-7)}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet))
 console.log('UNMERGE TASK 1 (ONE EDITOR SESSION)',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.wholeTaskVerification?.ok,true);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,4);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,api=>retryInvocation(api,sheet))
 console.log('UNMERGE TASK 2 IDEMPOTENT REINVOCATION',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.wholeTaskVerification?.ok,true);assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT UNMERGE LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
