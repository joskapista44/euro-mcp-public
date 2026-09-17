'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const batch=require('./xlsx-persistent-batch.cjs')
process.env.EURO_PLAYWRIGHT_PATH ||= '/home/user/marveen/node_modules/playwright'
function task(sheet,name){return {operations:[
 {intent:'sort_range',sheet,range:'A1:C5',keyRange:'A1:A5',order:'asc'},
 {intent:'format_range',sheet,range:'A1:C1',format:{bold:true,fillColor:[210,220,230]}},
 {intent:'layout_range',sheet,range:'A1:A5',type:'column.width',width:24},
 {intent:'freeze_panes',sheet,mode:'at',range:'A2'},
 {intent:'filter_range',sheet,range:'A1:C5',field:3,criteria1:'A',operator:'xlOr'},
 {intent:'set_validation',sheet,range:'E2:E5',validationType:'xlValidateWholeNumber',formula1:'1',formula2:'10'},
 {intent:'set_defined_name',name,refersTo:'='+sheet+'!$A$1:$C$5'}
]}}
function report(r){return {ok:r.ok,outcome:r.outcome,noOp:r.noOp,session:r.persistentSession,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:r}}
;(async()=>{
 const vault=require('/home/user/marveen/dist/web/vault.js')
 const pass=vault.getSecret('Elliot_nc_pass','xlsx-batch-live-acceptance');if(!pass)throw Error('vault secret missing')
 const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass,fileId:Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187),timeoutMs:30000,pollMs:50}
 const suffix=Date.now(),sheet='EURO_BATCH_'+suffix,name='EURO_BN_'+suffix,t=task(sheet,name)
 const first=await persistent.withPersistentXlsxSession(options,async api=>{
  const created=await api.createSheetVerified(sheet);if(!created.ok)return created
  const written=await api.writeRangeVerified({sheet,range:'A1:C5',values:[['Name','Score','Group'],['Delta',40,'B'],['Alpha',10,'A'],['Charlie',30,'A'],['Bravo',20,'B']]});if(!written.ok)return written
  return batch.executeBatchTask({task:t,api})
 })
 console.log('BATCH APPLY',JSON.stringify(report(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false)
 assert.equal(first.wholeTaskVerification.checks.length,7);assert.equal(first.wholeTaskVerification.readOnly,true)
 assert.equal(first.persistentSession.oneEditorSession,true);assert.equal(first.persistentSession.writes,9);assert.equal(first.persistentSession.persistenceBarrier.ok,true)
 const retry=await batch.executeBatchTaskInPersistentSession({...options,task:t})
 console.log('BATCH PERSISTED RETRY',JSON.stringify(report(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true)
 assert.equal(retry.persistentSession.oneEditorSession,true);assert.equal(retry.persistentSession.writes,0);assert.equal(retry.persistentSession.persistenceBarrier,null)
 console.log('XLSX CROSS-CAPABILITY PERSISTENT BATCH LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e.stack||e);process.exitCode=1})
