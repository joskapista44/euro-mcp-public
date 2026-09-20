'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-layout-batch-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet=`EURO PAGE ${String(Date.now()).slice(-7)}`
 const values=[['EURO SALES PRINT REPORT','Q1','Q2','Q3'],['North',120,135,150],['West',98,112,127],['South',87,101,119],['Central',110,125,142]]
 const task={operations:[
  {intent:'create_sheet',name:sheet},
  {intent:'write_range',sheet,range:'A1:D5',values,formulas:Array.from({length:5},()=>Array(4).fill(null))},
  {intent:'format_range',sheet,range:'A1:D1',format:{bold:true}},
  {intent:'layout_range',sheet,range:'A:D',type:'column.width',width:20},
  {intent:'set_page_layout',sheet,orientation:'xlLandscape',topMargin:10,bottomMargin:10,leftMargin:8,rightMargin:8,printGridlines:false,printHeadings:false}
 ],readbacks:[{sheet,range:'A1:D5'}]}
 const first=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('PAGE LAYOUT BATCH FIRST',JSON.stringify({ok:first.ok,outcome:first.outcome,authority:first.authority,noOp:first.noOp,writes:first.persistentSession?.writes,barrier:first.persistentSession?.persistenceBarrier,checks:first.wholeTaskVerification?.checks,readbacks:first.wholeTaskVerification?.readbacks,steps:first.steps?.map(x=>({index:x.index,intent:x.intent,outcome:x.result?.outcome,noOp:x.result?.noOp,state:x.result?.wholeTaskVerification?.state}))},null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.ok(first.persistentSession?.writes>0);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(first.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(first.wholeTaskVerification?.readbacks?.length,1)
 const retry=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('PAGE LAYOUT BATCH RETRY',JSON.stringify({ok:retry.ok,outcome:retry.outcome,authority:retry.authority,noOp:retry.noOp,writes:retry.persistentSession?.writes,barrier:retry.persistentSession?.persistenceBarrier,checks:retry.wholeTaskVerification?.checks,readbacks:retry.wholeTaskVerification?.readbacks,steps:retry.steps?.map(x=>({index:x.index,intent:x.intent,outcome:x.result?.outcome,noOp:x.result?.noOp,state:x.result?.wholeTaskVerification?.state}))},null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(retry.wholeTaskVerification?.readbacks?.length,1)
 const ps=retry.steps.find(x=>x.intent==='set_page_layout')?.result?.wholeTaskVerification?.state;assert.equal(ps.orientation,'xlLandscape');assert.ok(Math.abs(ps.topMargin-10)<.01);assert.ok(Math.abs(ps.leftMargin-8)<.01);assert.equal(ps.printGridlines,false);assert.equal(ps.printHeadings,false)
 console.log('XLSX PAGE LAYOUT BATCH LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
