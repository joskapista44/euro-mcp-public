'use strict'
const assert=require('assert/strict')
const batch=require('./xlsx-persistent-batch.cjs')
const {buildCrossFamilyTask,withRetryReceipts}=require('./xlsx-cross-family-agent-task.cjs')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js'),r=v.getSecret(id,'xlsx-cross-family-agent-task-live');if(!r)throw new Error('vault secret not found: '+id);return r}
function liveOptions(){return {url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}}
function readback(result,sheet){return result.wholeTaskVerification?.readbacks?.find(item=>item.sheet===sheet)}
function cell(read,address){for(const row of read?.cells||[])for(const item of row||[])if(item.address===address)return item;return null}
function scalar(item){for(const value of [item?.rawValue,item?.value])if(value!==null&&value!==undefined&&value!==''&&Number.isFinite(Number(value)))return Number(value);return null}
function value(item){return item?.rawValue??item?.value??null}
function assertBlank(item){assert(item,'blank cell must be present in LIVE readback');assert.equal(item.dataType,'blank');assert.ok(item.formula==null||item.formula==='');assert.ok([null,undefined,''].includes(value(item)))}
function assertRetryReceipts(task,result){
 for(const [index,op] of task.operations.entries()){
  if(!['insert_rows','move_range'].includes(op.intent)&&!['columns.autofit','rows.autofit'].includes(op.type))continue
  const step=result.steps?.find(step=>step.index===index&&step.intent===op.intent)
  assert.ok(step?.result?.retryToken,`missing operation-bound receipt: ${op.intent} at index ${index}`)
 }
}
function summary(result){const failed=result.ok?null:result.steps?.find(step=>step.result?.ok!==true)?.result;return {ok:result.ok,outcome:result.outcome,authority:result.authority,noOp:result.noOp,session:result.persistentSession,checks:result.wholeTaskVerification?.checks?.length,readbacks:result.wholeTaskVerification?.readbacks?.length,steps:result.steps?.map(step=>({index:step.index,intent:step.intent,outcome:step.result?.outcome,noOp:step.result?.noOp,retryReceipt:!!step.result?.retryToken})),diagnostic:failed&&{outcome:failed.outcome,index:failed.index,receipt:failed.receipt,result:failed.result,wholeTaskVerification:failed.wholeTaskVerification}}}
function assertFinalState(result,contract){
 const input=readback(result,contract.names.input),plan=readback(result,contract.names.plan),staging=readback(result,contract.names.staging),report=readback(result,contract.names.report)
 assert(input&&plan&&staging&&report,'all four final LIVE readbacks are required')
 assert.equal(value(cell(input,'A2')),contract.expected.topAccount);assert.equal(scalar(cell(input,'C2')),contract.expected.topActual)
 assert.equal(value(cell(report,'A2')),contract.expected.topAccount);assert.equal(scalar(cell(report,'C2')),contract.expected.topActual)
 assert.equal(value(cell(report,'G2')),contract.expected.movedControl);assertBlank(cell(staging,'A1'))
 assertBlank(cell(plan,'A2'));assert.equal(value(cell(plan,'A3')),'North follow-up')
}
async function main(){
 const options=liveOptions()
 const runId=String(Date.now()).slice(-7),task=buildCrossFamilyTask(runId)
 const first=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('CROSS-FAMILY APPLY',JSON.stringify(summary(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false)
 assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.closedByWrapper,true);assert.ok(first.persistentSession?.writes>0);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 assert.ok(Number.isFinite(first.persistentSession?.openedMs));assert.ok(Number.isFinite(first.persistentSession?.taskMs))
 assert.equal(first.wholeTaskVerification?.readOnly,true);assert.equal(first.wholeTaskVerification?.checks?.length,task.expected.operationCount);assert(first.wholeTaskVerification.checks.every(check=>check.ok));assert.equal(first.wholeTaskVerification.readbacks.length,4)
 assertRetryReceipts(task,first)
 assertFinalState(first,task)
 const retryTask=withRetryReceipts(task,first),retry=await batch.executeBatchTaskInPersistentSession({...options,task:retryTask})
 console.log('CROSS-FAMILY PERSISTED REOPEN RETRY',JSON.stringify(summary(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true)
 assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.closedByWrapper,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 assert.ok(Number.isFinite(retry.persistentSession?.openedMs));assert.ok(Number.isFinite(retry.persistentSession?.taskMs))
 assert.equal(retry.wholeTaskVerification?.readOnly,true);assert.equal(retry.wholeTaskVerification?.checks?.length,task.expected.operationCount);assert(retry.wholeTaskVerification.checks.every(check=>check.ok));assert.equal(retry.wholeTaskVerification.readbacks.length,4)
 assertFinalState(retry,task)
 console.log('XLSX CROSS-FAMILY COMPLEX AGENT-TASK LIVE ACCEPTANCE: PASS')
}
module.exports={assertRetryReceipts,assertBlank,assertFinalState}
if(require.main===module)main().catch(error=>{console.error(error?.stack||error);process.exitCode=1})
