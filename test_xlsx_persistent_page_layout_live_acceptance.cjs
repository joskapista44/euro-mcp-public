'use strict'
const assert=require('assert/strict')
const persistent=require('./xlsx-persistent-session.cjs'),layout=require('./xlsx-persistent-page-layout.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-page-layout-live-acceptance');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const suffix=String(Date.now()).slice(-7),sheet=`EURO PRINT ${suffix}`
 const setup=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);if(!c.ok)return c;const w=await api.writeRangeVerified({sheet,range:'A1:D5',values:[['Print','Layout','Acceptance','M7'],[1,2,3,4],[5,6,7,8],[9,10,11,12],[13,14,15,16]]});return w.ok?{ok:true,outcome:'page-layout-fixture-ready',authority:'LIVE_VERIFY',noOp:false}:{...w}})
 assert.equal(setup.ok,true)
 const task={operations:[{intent:'set_page_layout',sheet,orientation:'xlLandscape',topMargin:10,bottomMargin:11,leftMargin:12,rightMargin:13,printGridlines:true,printHeadings:true}]}
 const first=await layout.executePageLayoutTaskInPersistentSession({...options,task})
 console.log('PAGE LAYOUT FIRST',JSON.stringify(first,null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(first.persistentSession?.writes,1);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const retry=await layout.executePageLayoutTaskInPersistentSession({...options,task})
 console.log('PAGE LAYOUT RETRY',JSON.stringify(retry,null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 const s=retry.wholeTaskVerification?.state;assert.equal(s.orientation,'xlLandscape');assert.ok(Math.abs(s.topMargin-10)<.01);assert.ok(Math.abs(s.bottomMargin-11)<.01);assert.ok(Math.abs(s.leftMargin-12)<.01);assert.ok(Math.abs(s.rightMargin-13)<.01);assert.equal(s.printGridlines,true);assert.equal(s.printHeadings,true)
 console.log('XLSX PERSISTENT PAGE LAYOUT LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
