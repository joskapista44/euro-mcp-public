'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const nameAgent=require('./xlsx-agent-defined-name-task.cjs')
const namePersistent=require('./xlsx-persistent-defined-name.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-defined-name-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,definedNameObserved:async(spec,apply)=>{const r=await namePersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function setTask(name,refersTo){return {operations:[{intent:'set_defined_name',name,refersTo}]}}
function renameTask(name,newName,refersTo){return {operations:[{intent:'rename_defined_name',name,newName,refersTo}]}}
function deleteTask(name){return {operations:[{intent:'delete_defined_name',name}]}}
function exactSet(r,name,refersTo){const s=r?.wholeTaskVerification?.state;return s?.present===true&&s?.name===name&&s?.refersTo===refersTo}
function exactRename(r,name,newName,refersTo){const s=r?.wholeTaskVerification?.state;return s?.source?.present===false&&s?.source?.name===name&&s?.target?.present===true&&s?.target?.name===newName&&s?.target?.refersTo===refersTo}
function exactDelete(r,name){const s=r?.wholeTaskVerification?.state;return s?.present===false&&s?.name===name}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
async function run(task){return persistent.withPersistentXlsxSession(options,api=>nameAgent.executeDefinedNameTask({task,api:liveApi(api)}))}
function assertMutation(r){assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.noOp,false);assert.equal(r.persistentSession?.oneEditorSession,true);assert.equal(r.persistentSession?.writes,1);assert.equal(r.persistentSession?.persistenceBarrier?.ok,true)}
function assertRetry(r){assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.noOp,true);assert.equal(r.persistentSession?.oneEditorSession,true);assert.equal(r.persistentSession?.writes,0);assert.equal(r.persistentSession?.persistenceBarrier,null)}
;(async()=>{
 const suffix=String(Date.now()).slice(-7),name=`EURO_N_${suffix}`,renamed=`EURO_R_${suffix}`,refersTo='=Sheet1!$XFD$30'
 const set=await run(setTask(name,refersTo));console.log('DEFINED NAME SET TASK',JSON.stringify(d(set),null,2));assertMutation(set);assert.equal(exactSet(set,name,refersTo),true)
 const setRetry=await run(setTask(name,refersTo));console.log('DEFINED NAME SET RETRY',JSON.stringify(d(setRetry),null,2));assertRetry(setRetry);assert.equal(exactSet(setRetry,name,refersTo),true)
 const renamedResult=await run(renameTask(name,renamed,refersTo));console.log('DEFINED NAME RENAME TASK',JSON.stringify(d(renamedResult),null,2));assertMutation(renamedResult);assert.equal(exactRename(renamedResult,name,renamed,refersTo),true)
 const renameRetry=await run(renameTask(name,renamed,refersTo));console.log('DEFINED NAME RENAME RETRY',JSON.stringify(d(renameRetry),null,2));assertRetry(renameRetry);assert.equal(exactRename(renameRetry,name,renamed,refersTo),true)
 const deleted=await run(deleteTask(renamed));console.log('DEFINED NAME DELETE TASK',JSON.stringify(d(deleted),null,2));assertMutation(deleted);assert.equal(exactDelete(deleted,renamed),true)
 const deleteRetry=await run(deleteTask(renamed));console.log('DEFINED NAME DELETE RETRY',JSON.stringify(d(deleteRetry),null,2));assertRetry(deleteRetry);assert.equal(exactDelete(deleteRetry,renamed),true)
 console.log('XLSX PERSISTENT DEFINED NAME LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
