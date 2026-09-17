'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-chart-task.cjs')
const chart=require('./xlsx-persistent-chart.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-chart-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,chartObserved:async(op,apply)=>{const r=await chart.chartObserved(api.session,op,apply);if(apply&&(r?.applied||r?.mutation?.ok))api.session.markWrite();return r}}}
function setTask(sheet,name){return {operations:[{intent:'set_chart',sheet,name,range:'A1:B3',chartType:'bar',title:'Revenue',width:3600000,height:2160000,expectedSeriesCount:1}]}}
function deleteTask(sheet,name){return {operations:[{intent:'delete_chart',sheet,name}]}}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
;(async()=>{
 const suffix=String(Date.now()).slice(-7),sheet=`EURO_CH_${suffix}`,name=`EURO_C_${suffix}`
 const first=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);if(!c.ok)return c;const w=await api.writeRangeVerified({sheet,range:'A1:B3',values:[['Month','Revenue'],['Jan',10],['Feb',20]]});if(!w.ok)return w;return agent.executeChartTask({task:setTask(sheet,name),api:liveApi(api)})});console.log('CHART SET TASK',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,api=>agent.executeChartTask({task:setTask(sheet,name),api:liveApi(api)}));console.log('CHART SET RETRY',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 const deleted=await persistent.withPersistentXlsxSession(options,api=>agent.executeChartTask({task:deleteTask(sheet,name),api:liveApi(api)}));console.log('CHART DELETE TASK',JSON.stringify(d(deleted),null,2))
 assert.equal(deleted.ok,true);assert.equal(deleted.noOp,false);assert.equal(deleted.persistentSession?.writes,1);assert.equal(deleted.persistentSession?.persistenceBarrier?.ok,true)
 const deleteRetry=await persistent.withPersistentXlsxSession(options,api=>agent.executeChartTask({task:deleteTask(sheet,name),api:liveApi(api)}));console.log('CHART DELETE RETRY',JSON.stringify(d(deleteRetry),null,2))
 assert.equal(deleteRetry.ok,true);assert.equal(deleteRetry.noOp,true);assert.equal(deleteRetry.persistentSession?.writes,0);assert.equal(deleteRetry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT CHART LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
