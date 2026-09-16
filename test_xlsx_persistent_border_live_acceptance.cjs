'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const agent=require('./xlsx-agent-format-task.cjs')
const format=require('./xlsx-persistent-format.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-border-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const sessionOptions={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function compact(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,steps:r.steps?.map(x=>({outcome:x.outcome,noOp:x.noOp,verification:x.wholeTaskVerification}))}}
;(async()=>{
 const sheet=`EURO BRD ${String(Date.now()).slice(-7)}`
 const setup=await persistent.executeAgentTaskInPersistentSession({fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:K5',values:Array.from({length:5},(_,r)=>Array.from({length:11},(_,c)=>r*11+c+1))}]}})
 assert.equal(setup.ok,true)
 const cases=[
  ['Top','A1:B2',[10,20,30]],['Bottom','D1:E2',[20,30,40]],['Left','G1:H2',[30,40,50]],['Right','J1:K2',[40,50,60]],
  ['InsideHorizontal','A4:B5',[50,60,70]],['InsideVertical','D4:E5',[60,70,80]],['DiagonalDown','G4:H5',[70,80,90]],['DiagonalUp','J4:K5',[80,90,100]]
 ]
 const first=await persistent.withPersistentXlsxSession(sessionOptions,async api=>{
  const adapter={...api,formatRangeObserved:async spec=>{const x=await format.runCommand(api.session,[spec.sheet,spec.range,spec.format,!!spec.apply]);if(spec.apply&&x?.ok&&!x.noOp)api.session.markWrite();return x}}
  const steps=[]
  for(const [index,range,color] of cases){const task={operations:[{intent:'format_range',sheet,range,format:{border:{index,style:'Thin',color}}}]};const r=await agent.executeFormatTask({task,api:adapter});steps.push(r);if(!r.ok)return {ok:false,outcome:'xlsx-border-integrated-step-failed',authority:r.authority||'LIVE_READ',noOp:false,steps}}
  return {ok:true,outcome:'xlsx-border-integrated-live-verified',authority:'LIVE_VERIFY',noOp:false,steps}
 })
 console.log('BORDER FIRST',JSON.stringify(compact(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.persistentSession?.writes,8);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const [index,range,color]=cases[cases.length-1]
 const retry=await format.executeFormatTaskInPersistentSession({fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50,task:{operations:[{intent:'format_range',sheet,range,format:{border:{index,style:'Thin',color}}}]}})
 console.log('BORDER RETRY',JSON.stringify(compact(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 console.log('XLSX PERSISTENT BORDER LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
