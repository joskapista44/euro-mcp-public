'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const formatAgent=require('./xlsx-agent-format-task.cjs'),formatPersistent=require('./xlsx-persistent-format.cjs')
const layoutAgent=require('./xlsx-agent-layout-task.cjs'),layoutPersistent=require('./xlsx-persistent-layout.cjs')
const mergeAgent=require('./xlsx-agent-merge-task.cjs'),mergePersistent=require('./xlsx-persistent-merge.cjs')
const cfAgent=require('./xlsx-agent-conditional-format-task.cjs'),cfPersistent=require('./xlsx-persistent-conditional-format.cjs')
const freezeAgent=require('./xlsx-agent-freeze-task.cjs'),freezePersistent=require('./xlsx-persistent-freeze.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-m4-integrated-persistent-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function adapters(api){
 function written(r,apply){if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}
 return {
  format:{...api,formatRangeObserved:async spec=>written(await formatPersistent.runCommand(api.session,[spec.sheet,spec.range,spec.format,!!spec.apply]),!!spec.apply)},
  layout:{...api,layoutObserved:async(spec,apply)=>written(await layoutPersistent.runCommand(api.session,spec,apply),apply)},
  merge:{...api,mergeObserved:async(spec,apply)=>written(await mergePersistent.runCommand(api.session,spec,apply),apply)},
  cf:{...api,cfObserved:async spec=>{const r=await cfPersistent.runCommand(api.session,spec),out={...r,authority:cfPersistent.authorityFor(spec,r)};if(spec.type!=='cf.inspect'&&out.ok&&out.authority==='LIVE_VERIFY')api.session.markWrite();return out}},
  freeze:{...api,freezeObserved:async(op,apply)=>written(await freezePersistent.runCommand(api.session,op,apply),apply)}
 }
}
function operations(sheet){
 return [
  {name:'format-extended',agent:formatAgent.executeFormatTask,api:'format',task:{operations:[{intent:'format_range',sheet,range:'A1:B2',format:{bold:true,fontColor:[10,20,30],fillColor:[210,220,230],alignHorizontal:'center',alignVertical:'top',wrap:true,numberFormat:'0.00'}}]}},
  {name:'border',agent:formatAgent.executeFormatTask,api:'format',task:{operations:[{intent:'format_range',sheet,range:'A1:B2',format:{border:{index:'Top',style:'Thin',color:[30,60,90]}}}]}},
  {name:'layout',agent:layoutAgent.executeLayoutTask,api:'layout',task:{operations:[{intent:'layout_range',sheet,range:'D1:D3',type:'column.width',width:24}]}},
  {name:'merge',agent:mergeAgent.executeMergeTask,api:'merge',task:{operations:[{intent:'merge_range',sheet,range:'G1:H2'}]}},
  {name:'conditional-format',agent:cfAgent.executeConditionalFormatTask,api:'cf',task:{operations:[{intent:'add_conditional_format',sheet,range:'J1:J3',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:[255,0,0],priority:1}}]}},
  {name:'freeze',agent:freezeAgent.executeFreezeTask,api:'freeze',task:{operations:[{intent:'freeze_panes',sheet,mode:'at',range:'C4'}]}}
 ]
}
async function verifyWholeTask(a,sheet){
 const checks=[]
 for(const op of operations(sheet)){
  const result=await op.agent({task:op.task,api:a[op.api]});checks.push({name:op.name,result})
  if(!result.ok||result.authority!=='LIVE_VERIFY'||result.noOp!==true)return {ok:false,outcome:'m4-whole-task-live-verify-failed',authority:result.authority||'LIVE_READ',failedCheck:op.name,checks}
 }
 return {ok:true,outcome:'m4-whole-task-live-verified',authority:'LIVE_VERIFY',checks}
}
async function executeAgentTask(api,sheet,setup){
 const a=adapters(api),steps=[]
 if(setup){
  const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'m4-setup-create-failed',authority:created.authority||'LIVE_READ',noOp:false,steps:[{name:'setup-create',result:created}]}
  const values=[
   [1,2,'',10,'','', 'MERGE','','',5],
   [3,4,'',20,'','', '','','',15],
   ['','','',30,'','', '','','',25],
   ['','','','','','', '','','','']
  ]
  const written=await api.writeRangeVerified({sheet,range:'A1:J4',values});if(!written.ok)return {ok:false,outcome:'m4-setup-write-failed',authority:written.authority||'LIVE_READ',noOp:false,steps:[{name:'setup-write',result:written}]}
 }
 for(const op of operations(sheet)){
  const result=await op.agent({task:op.task,api:a[op.api]});steps.push({name:op.name,result})
  if(!result.ok)return {ok:false,outcome:'m4-integrated-step-failed',authority:result.authority||'LIVE_READ',noOp:false,failedStep:op.name,steps}
 }
 const wholeTaskVerification=await verifyWholeTask(a,sheet)
 if(!wholeTaskVerification.ok)return {ok:false,outcome:'m4-integrated-whole-task-verify-failed',authority:'PRIMITIVE_LIVE_VERIFY_ONLY',noOp:false,steps,wholeTaskVerification}
 const allNoOp=steps.every(x=>x.result.noOp===true)
 return {ok:true,outcome:allNoOp?'m4-integrated-already-satisfied':'m4-integrated-live-verified',authority:'LIVE_VERIFY',noOp:allNoOp,steps,wholeTaskVerification}
}
function compact(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,failedStep:r.failedStep,wholeTaskVerification:r.wholeTaskVerification&&{ok:r.wholeTaskVerification.ok,outcome:r.wholeTaskVerification.outcome,failedCheck:r.wholeTaskVerification.failedCheck,checks:r.wholeTaskVerification.checks?.map(x=>({name:x.name,ok:x.result.ok,authority:x.result.authority,noOp:x.result.noOp}))},steps:r.steps?.map(x=>({name:x.name,ok:x.result.ok,outcome:x.result.outcome,authority:x.result.authority,noOp:x.result.noOp}))}}
;(async()=>{
 const sheet=`EURO M4 ${String(Date.now()).slice(-7)}`
 const first=await persistent.withPersistentXlsxSession(options,api=>executeAgentTask(api,sheet,true))
 console.log('M4 INTEGRATED FIRST',JSON.stringify(compact(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.wholeTaskVerification?.ok,true);assert.equal(first.persistentSession?.writes,8);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await persistent.withPersistentXlsxSession(options,api=>executeAgentTask(api,sheet,false))
 console.log('M4 INTEGRATED RETRY',JSON.stringify(compact(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.wholeTaskVerification?.ok,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX M4 INTEGRATED PERSISTENT LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
