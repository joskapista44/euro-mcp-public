'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-freeze-task.cjs')
const freeze=require('./xlsx-persistent-freeze.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-freeze-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const sessionOptions={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function compact(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,steps:r.steps?.map(x=>({outcome:x.outcome,noOp:x.noOp,verification:x.wholeTaskVerification?.verification}))}}
;(async()=>{
  const sheet=`EURO FRZ ${String(Date.now()).slice(-7)}`
  const setup=await persistent.executeAgentTaskInPersistentSession({fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:D5',values:[[1,2,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16],[17,18,19,20]]}]}})
  assert.equal(setup.ok,true)
  const first=await persistent.withPersistentXlsxSession(sessionOptions,async api=>{
    const adapter={...api,freezeObserved:async(op,apply)=>{const x=await freeze.runCommand(api.session,op,apply);if(apply&&x?.ok&&x.applied===true)api.session.markWrite();return x}}
    const tasks=[
      {operations:[{intent:'freeze_panes',sheet,mode:'rows',count:2}]},
      {operations:[{intent:'freeze_panes',sheet,mode:'columns',count:2}]},
      {operations:[{intent:'unfreeze_panes',sheet}]},
      {operations:[{intent:'freeze_panes',sheet,mode:'at',range:'C4'}]}
    ]
    const steps=[]
    for(const task of tasks){const r=await agent.executeFreezeTask({task,api:adapter});steps.push(r);if(!r.ok)return {ok:false,outcome:'xlsx-freeze-integrated-step-failed',authority:r.authority||'LIVE_READ',noOp:false,steps}}
    return {ok:true,outcome:'xlsx-freeze-integrated-live-verified',authority:'LIVE_VERIFY',noOp:false,steps}
  })
  console.log('FREEZE FIRST',JSON.stringify(compact(first),null,2))
  assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.persistentSession?.writes,4);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
  const retryTask={operations:[{intent:'freeze_panes',sheet,mode:'at',range:'C4'}]}
  const retry=await freeze.executeFreezeTaskInPersistentSession({fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50,task:retryTask})
  console.log('FREEZE RETRY',JSON.stringify(compact(retry),null,2))
  assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
  console.log('XLSX PERSISTENT FREEZE LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
