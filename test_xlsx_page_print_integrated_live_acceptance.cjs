'use strict'
const assert=require('assert/strict'),batch=require('./xlsx-persistent-batch.cjs')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-print-integrated-live');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet='EURO PAGE PRINT '+String(Date.now()).slice(-7)
 const values=[['EURO MANAGEMENT REPORT','Q1','Q2','Q3','Q4','Target','Variance','Status'],...Array.from({length:23},(_,i)=>['Region '+(i+1),100+i,110+i,120+i,130+i,450+i*4,10,'On track'])]
 const formulas=Array.from({length:24},()=>Array(8).fill(null))
 const task={operations:[
 {intent:'create_sheet',name:sheet},
 {intent:'write_range',sheet,range:'A1:H24',values,formulas},
 {intent:'format_range',sheet,range:'A1:H1',format:{bold:true}},
 {intent:'layout_range',sheet,range:'A1:H24',type:'column.width',width:18},
 {intent:'set_page_layout',sheet,orientation:'xlLandscape',topMargin:10,bottomMargin:10,leftMargin:8,rightMargin:8,printGridlines:false,printHeadings:false},
 {intent:'set_print_setup',sheet,mode:'page_size',width:210,height:297},
 {intent:'set_print_setup',sheet,mode:'fit_to_pages',fitToWidth:1,fitToHeight:0},
 {intent:'set_print_titles',sheet,axis:'rows',from:1,to:1},
 {intent:'set_print_titles',sheet,axis:'columns',from:1,to:1},
 {intent:'set_print_area',sheet,mode:'set',range:'A1:H24'},
 {intent:'set_header_footer',sheet,slot:'oddHeader',value:'&LEURO REPORT&C&P / &N'},
 {intent:'set_header_footer',sheet,slot:'oddFooter',value:'&LConfidential&R&D'},
 {intent:'set_page_break',sheet,mode:'add',axis:'row',at:13},
 {intent:'set_first_page_number',sheet,value:3}
 ],readbacks:[{sheet,range:'A1:H24'}]}
 const first=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('PAGE PRINT INTEGRATED FIRST',JSON.stringify({ok:first.ok,outcome:first.outcome,authority:first.authority,noOp:first.noOp,writes:first.persistentSession?.writes,barrier:first.persistentSession?.persistenceBarrier,checks:first.wholeTaskVerification?.checks,readbacks:first.wholeTaskVerification?.readbacks?.length,steps:first.steps?.map(s=>({index:s.index,intent:s.intent,outcome:s.result?.outcome,noOp:s.result?.noOp}))},null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true);assert.equal(first.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(first.wholeTaskVerification?.readbacks?.length,1)
 const retry=await batch.executeBatchTaskInPersistentSession({...options,task})
 console.log('PAGE PRINT INTEGRATED RETRY',JSON.stringify({ok:retry.ok,outcome:retry.outcome,authority:retry.authority,noOp:retry.noOp,writes:retry.persistentSession?.writes,barrier:retry.persistentSession?.persistenceBarrier,checks:retry.wholeTaskVerification?.checks,readbacks:retry.wholeTaskVerification?.readbacks?.length,steps:retry.steps?.map(s=>({index:s.index,intent:s.intent,outcome:s.result?.outcome,noOp:s.result?.noOp}))},null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.wholeTaskVerification?.checks?.every(x=>x.ok),true);assert.equal(retry.wholeTaskVerification?.readbacks?.length,1)
 console.log('XLSX PAGE PRINT INTEGRATED LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
