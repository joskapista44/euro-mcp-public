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
  // Multi-sheet discriminator for the full-batch failure: three sequential
  // FreezeAt operations on three different sheets in ONE editor session.
  const sheets=[`EURO FRZA ${String(Date.now()).slice(-7)}`,`EURO FRZB ${String(Date.now()).slice(-7)}`,`EURO FRZC ${String(Date.now()).slice(-7)}`]
  const multi=await persistent.withPersistentXlsxSession(sessionOptions,async api=>{
    const steps=[]
    for(const s of sheets){const cr=await api.createSheetVerified(s);if(!cr.ok)return {ok:false,outcome:'freeze-multi-sheet-create-failed',steps};const wr=await api.writeRangeVerified({sheet:s,range:'A1:D5',values:[[1,2,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16],[17,18,19,20]]});if(!wr.ok)return {ok:false,outcome:'freeze-multi-sheet-write-failed',steps}}
    const adapter={...api,freezeObserved:async(op,apply)=>{const x=await freeze.runCommand(api.session,op,apply);if(apply&&x?.ok&&x.applied===true)api.session.markWrite();return x}}
    for(let i=0;i<sheets.length;i++){const r=await agent.executeFreezeTask({task:{operations:[{intent:'freeze_panes',sheet:sheets[i],mode:'at',range:i===0?'A3':'A4'}]},api:adapter});steps.push(r);if(!r.ok)return {ok:false,outcome:'freeze-multi-sheet-step-failed',authority:r.authority||'LIVE_READ',steps}}
    return {ok:true,outcome:'freeze-multi-sheet-live-verified',authority:'LIVE_VERIFY',steps}
  })
  console.log('FREEZE MULTI-SHEET',JSON.stringify(compact(multi),null,2))
  assert.equal(multi.ok,true);assert.equal(multi.authority,'LIVE_VERIFY');assert.equal(multi.steps.length,3);assert.equal(multi.steps.every(x=>x.ok&&x.authority==='LIVE_VERIFY'),true);assert.equal(multi.persistentSession?.writes,9);assert.equal(multi.persistentSession?.persistenceBarrier,null)
  // Monthly_Plan-like discriminator: formula-heavy + formatted/layout sheet,
  // conditional formatting, then sequential freeze after another sheet froze.
  const ctxA=`EURO FCA ${String(Date.now()).slice(-7)}`,ctxB=`EURO FCB ${String(Date.now()).slice(-7)}`
  const context=await persistent.withPersistentXlsxSession(sessionOptions,async api=>{
    for(const s of [ctxA,ctxB]){const cr=await api.createSheetVerified(s);if(!cr.ok)return cr}
    const wa=await api.writeRangeVerified({sheet:ctxA,range:'A1:D5',values:[[1,2,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16],[17,18,19,20]]});if(!wa.ok)return wa
    const rows=Array.from({length:15},()=>Array(8).fill('')),forms=Array.from({length:15},()=>Array(8).fill(''))
    rows[0][0]='MONTHLY PLAN VS ACTUAL 2026';['Month','Actual Revenue','Revenue Target','Variance','Attainment','Actual Profit','Profit Target','Profit Variance'].forEach((h,j)=>rows[2][j]=h)
    for(let i=0;i<12;i++){const row=i+4;rows[row-1][0]='M'+(i+1);rows[row-1][1]=(i+1)*1000;rows[row-1][2]=(i+1)*950;rows[row-1][5]=(i+1)*200;rows[row-1][6]=(i+1)*180;forms[row-1][3]='=B'+row+'-C'+row;forms[row-1][4]='=B'+row+'/C'+row;forms[row-1][7]='=F'+row+'-G'+row}
    const wb=await api.writeRangeVerified({sheet:ctxB,range:'A1:H15',values:rows,formulas:forms});if(!wb.ok)return {ok:false,outcome:'freeze-monthly-write-failed',authority:wb.authority||'LIVE_READ',written:wb}
    const adapter={...api,freezeObserved:async(op,apply)=>{const x=await freeze.runCommand(api.session,op,apply);if(apply&&x?.ok&&x.applied===true)api.session.markWrite();return x}}
    const a=await agent.executeFreezeTask({task:{operations:[{intent:'freeze_panes',sheet:ctxA,mode:'at',range:'A3'}]},api:adapter});if(!a.ok)return {ok:false,outcome:'freeze-context-first-failed',steps:[a]}
    const b=await agent.executeFreezeTask({task:{operations:[{intent:'freeze_panes',sheet:ctxB,mode:'at',range:'A4'}]},api:adapter});return b.ok?{ok:true,outcome:'freeze-monthly-context-live-verified',authority:'LIVE_VERIFY',steps:[a,b]}:{ok:false,outcome:'freeze-monthly-context-second-failed',authority:b.authority||'LIVE_READ',steps:[a,b]}
  })
  console.log('FREEZE MONTHLY-CONTEXT',JSON.stringify(compact(context),null,2))
  assert.equal(context.ok,true);assert.equal(context.authority,'LIVE_VERIFY');assert.equal(context.steps.length,2);assert.equal(context.steps.every(x=>x.ok&&x.authority==='LIVE_VERIFY'),true)
      console.log('XLSX PERSISTENT FREEZE LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
