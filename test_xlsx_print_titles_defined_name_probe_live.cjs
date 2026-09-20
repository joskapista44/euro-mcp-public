'use strict'
const assert=require('assert/strict'),persistent=require('./xlsx-persistent-session.cjs'),names=require('./xlsx-persistent-defined-name.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-print-titles-defined-name-probe');if(!r)throw new Error('vault secret not found: '+id);return r}
const options={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass'),fileId:FILE_ID,timeoutMs:30000,pollMs:50}
;(async()=>{
 const sheet=`EURO PTDN ${String(Date.now()).slice(-7)}`
 const fixture=await persistent.withPersistentXlsxSession(options,async api=>{const c=await api.createSheetVerified(sheet);return c.ok?{ok:true,outcome:'ptdn-fixture-ready',authority:'LIVE_VERIFY',noOp:false}:c});assert.equal(fixture.ok,true);assert.equal(fixture.persistentSession?.persistenceBarrier?.ok,true)
 const candidates=[{name:'Print_Titles',refersTo:`='${sheet}'!$1:$2`},{name:'_xlnm.Print_Titles',refersTo:`='${sheet}'!$1:$2`}]
 for(const spec of candidates){
  const result=await persistent.withPersistentXlsxSession(options,async api=>{const r=await names.runCommand(api.session,{intent:'set_defined_name',...spec},true);if(r?.ok&&!r.noOp)api.session.markWrite();return {...r,authority:r?.ok?'LIVE_VERIFY':'LIVE_READ'}})
  console.log('PRINT TITLES DEFINED NAME CANDIDATE',JSON.stringify({spec,result},null,2))
  if(result.ok){console.log('XLSX PRINT TITLES DEFINED NAME PROBE: PASS',spec.name);return}
 }
 throw new Error('no print-title defined-name candidate accepted')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
