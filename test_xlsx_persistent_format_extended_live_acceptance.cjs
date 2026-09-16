'use strict'
const assert=require('assert')
const util=require('util')
const {executeAgentTaskInPersistentSession}=require('./xlsx-persistent-session.cjs')
const {executeFormatTaskInPersistentSession}=require('./xlsx-persistent-format.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-format-extended-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const common={fileId:FILE_ID,credentials,timeoutMs:30000,pollMs:50}
function diagnostic(r){return {ok:r?.ok,outcome:r?.outcome,authority:r?.authority,writeAllowed:r?.writeAllowed,noOp:r?.noOp,writes:r?.persistentSession?.writes,barrier:r?.persistentSession?.persistenceBarrier,plan:r?.plan,pre:r?.pre,applied:r?.applied,final:r?.final,receipt:r?.receipt,wholeTaskVerification:r?.wholeTaskVerification,persistentSession:r?.persistentSession}}
function dump(label,r){console.log(label,util.inspect(diagnostic(r),{depth:null,colors:false,maxArrayLength:null,breakLength:120,compact:false}))}
function check(actual,expected,label,context){assert.deepStrictEqual(actual,expected,`${label}; diagnostic=${util.inspect(diagnostic(context),{depth:null,colors:false,breakLength:Infinity})}`)}
;(async()=>{
  const tag=String(Date.now()).slice(-7),sheet=`EURO FEX ${tag}`
  const setup=await executeAgentTaskInPersistentSession({...common,task:{operations:[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:B2',values:[[1,2],[3,4]]}]}})
  dump('FORMAT EXT SETUP',setup)
  check(setup.ok,true,'setup.ok',setup)
  const task={operations:[{intent:'format_range',sheet,range:'A1:B2',format:{fontColor:[10,20,30],fillColor:[210,220,230],alignVertical:'top'}}]}
  const first=await executeFormatTaskInPersistentSession({...common,task})
  dump('FORMAT EXT FIRST',first)
  check(first.ok,true,'first.ok',first)
  check(first.authority,'LIVE_VERIFY','first.authority',first)
  check(first.noOp,false,'first.noOp',first)
  check(first.persistentSession?.writes,1,'first.writes',first)
  check(first.persistentSession?.persistenceBarrier?.ok,true,'first.persistenceBarrier.ok',first)
  const retry=await executeFormatTaskInPersistentSession({...common,task})
  dump('FORMAT EXT RETRY',retry)
  check(retry.ok,true,'retry.ok',retry)
  check(retry.authority,'LIVE_VERIFY','retry.authority',retry)
  check(retry.noOp,true,'retry.noOp',retry)
  check(retry.persistentSession?.writes,0,'retry.writes',retry)
  check(retry.persistentSession?.persistenceBarrier,null,'retry.persistenceBarrier',retry)
  console.log('XLSX PERSISTENT FORMAT EXTENDED LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
