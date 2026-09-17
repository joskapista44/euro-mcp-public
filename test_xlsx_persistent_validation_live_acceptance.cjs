'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const validationAgent=require('./xlsx-agent-validation-task.cjs')
const validationPersistent=require('./xlsx-persistent-validation.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-validation-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,validationObserved:async(spec,apply)=>{const r=await validationPersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function setTask(sheet){return {operations:[{intent:'set_validation',sheet,range:'B2:B5',validationType:'xlValidateWholeNumber',alertStyle:'xlValidAlertStop',operator:'xlBetween',formula1:'1',formula2:'10',ignoreBlank:false,inCellDropdown:true,inputTitle:'M6.2 input',inputMessage:'Enter 1-10',showInput:true,showError:true,errorTitle:'M6.2 error',errorMessage:'Only 1-10'}]}}
function clearTask(sheet){return {operations:[{intent:'clear_validation',sheet,range:'B2:B5'}]}}
async function firstInvocation(api,sheet){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'validation-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const values=[['Item','Quantity'],['Alpha',1],['Bravo',2],['Charlie',3],['Delta',4]]
 const written=await api.writeRangeVerified({sheet,range:'A1:B5',values});if(!written.ok)return {ok:false,outcome:'validation-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 return validationAgent.executeValidationTask({task:setTask(sheet),api:liveApi(api)})
}
function exactSet(r){const s=r?.wholeTaskVerification?.state;return s?.address==='B2:B5'&&s?.present===true&&s?.type==='xlValidateWholeNumber'&&s?.alertStyle==='xlValidAlertStop'&&s?.operator==='xlBetween'&&s?.formula1==='1'&&s?.formula2==='10'&&s?.ignoreBlank===false&&s?.inCellDropdown===true&&s?.inputTitle==='M6.2 input'&&s?.inputMessage==='Enter 1-10'&&s?.showInput===true&&s?.showError===true&&s?.errorTitle==='M6.2 error'&&s?.errorMessage==='Only 1-10'}
function exactClear(r){const s=r?.wholeTaskVerification?.state;return s?.address==='B2:B5'&&s?.absent===true&&s?.present===false}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
;(async()=>{
 const sheet=`EURO VAL ${String(Date.now()).slice(-7)}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet))
 console.log('VALIDATION SET TASK 1 (ONE EDITOR SESSION)',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(exactSet(first),true);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const setRetry=await persistent.withPersistentXlsxSession(options,api=>validationAgent.executeValidationTask({task:setTask(sheet),api:liveApi(api)}))
 console.log('VALIDATION SET TASK 2 IDEMPOTENT REINVOCATION',JSON.stringify(d(setRetry),null,2))
 assert.equal(setRetry.ok,true);assert.equal(setRetry.authority,'LIVE_VERIFY');assert.equal(setRetry.noOp,true);assert.equal(exactSet(setRetry),true);assert.equal(setRetry.persistentSession?.oneEditorSession,true);assert.equal(setRetry.persistentSession?.writes,0);assert.equal(setRetry.persistentSession?.persistenceBarrier,null)
 const cleared=await persistent.withPersistentXlsxSession(options,api=>validationAgent.executeValidationTask({task:clearTask(sheet),api:liveApi(api)}))
 console.log('VALIDATION CLEAR TASK 3 (ONE EDITOR SESSION)',JSON.stringify(d(cleared),null,2))
 assert.equal(cleared.ok,true);assert.equal(cleared.authority,'LIVE_VERIFY');assert.equal(cleared.noOp,false);assert.equal(exactClear(cleared),true);assert.equal(cleared.persistentSession?.oneEditorSession,true);assert.equal(cleared.persistentSession?.writes,1);assert.equal(cleared.persistentSession?.persistenceBarrier?.ok,true)
 const clearRetry=await persistent.withPersistentXlsxSession(options,api=>validationAgent.executeValidationTask({task:clearTask(sheet),api:liveApi(api)}))
 console.log('VALIDATION CLEAR TASK 4 IDEMPOTENT REINVOCATION',JSON.stringify(d(clearRetry),null,2))
 assert.equal(clearRetry.ok,true);assert.equal(clearRetry.authority,'LIVE_VERIFY');assert.equal(clearRetry.noOp,true);assert.equal(exactClear(clearRetry),true);assert.equal(clearRetry.persistentSession?.oneEditorSession,true);assert.equal(clearRetry.persistentSession?.writes,0);assert.equal(clearRetry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT DATA VALIDATION LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
