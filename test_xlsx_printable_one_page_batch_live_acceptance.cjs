'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-printable-one-page-batch-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet=`EURO ONEPAGE ${String(Date.now()).slice(-7)}`
 const values=[['EUROPEAN SALES SUMMARY','Q1','Q2','Q3','Q4','Target','Variance','Status'],['North',120,135,150,162,550,17,'Above'],['West',98,112,127,139,460,16,'Above'],['South',87,101,119,130,430,7,'Above'],['Central',110,125,142,154,520,11,'Above'],['Total',415,473,538,585,1960,51,'Above']]
 const formulas=Array.from({length:6},()=>Array(8).fill(null))
 const task={operations:[
  {intent:'create_sheet',name:sheet},
  {intent:'write_range',sheet,range:'A1:H6',values,formulas},
  {intent:'format_range',sheet,range:'A1:H1',format:{bold:true}},
  {intent:'layout_range',sheet,range:'A1:H6',type:'column.width',width:18},
  {intent:'set_page_layout',sheet,orientation:'xlLandscape',topMargin:10,bottomMargin:10,leftMargin:8,rightMargin:8,printGridlines:false,printHeadings:false},
  {intent:'set_print_setup',sheet,mode:'page_size',width:210,height:297},
  {intent:'set_print_setup',sheet,mode:'fit_to_pages',fitToWidth:1,fitToHeight:1}
 ],readbacks:[{sheet,range:'A1:H6'}]}
 const first=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('ONE PAGE PRINT FIRST',JSON.stringify({ok:first.ok,outcome:first.outcome,authority:first.authority,noOp:first.noOp,writes:first.persistentSession?.writes,barrier:first.persistentSession?.persistenceBarrier,checks:first.wholeTaskVerification?.checks,readbacks:first.wholeTaskVerification?.readbacks?.length,steps:first.steps?.map(x=>({index:x.index,intent:x.intent,outcome:x.result?.outcome,noOp:x.result?.noOp,state:x.result?.wholeTaskVerification?.state}))},null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(first.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(first.wholeTaskVerification?.readbacks?.length,1)
 const retry=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('ONE PAGE PRINT RETRY',JSON.stringify({ok:retry.ok,outcome:retry.outcome,authority:retry.authority,noOp:retry.noOp,writes:retry.persistentSession?.writes,barrier:retry.persistentSession?.persistenceBarrier,checks:retry.wholeTaskVerification?.checks,readbacks:retry.wholeTaskVerification?.readbacks?.length,steps:retry.steps?.map(x=>({index:x.index,intent:x.intent,outcome:x.result?.outcome,noOp:x.result?.noOp,state:x.result?.wholeTaskVerification?.state}))},null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(retry.wholeTaskVerification?.readbacks?.length,1)
 const states=retry.steps.filter(x=>['set_page_layout','set_print_setup'].includes(x.intent)).map(x=>x.result?.wholeTaskVerification?.state).filter(Boolean)
 assert.ok(states.some(s=>s.orientation==='xlLandscape'&&Math.abs(s.topMargin-10)<.01&&Math.abs(s.leftMargin-8)<.01))
 assert.ok(states.some(s=>s.fitToWidth===1&&s.fitToHeight===1))
 assert.ok(states.some(s=>Math.abs(s.width-210)<.01&&Math.abs(s.height-297)<.01))
 console.log('XLSX PRINTABLE ONE PAGE BATCH LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
