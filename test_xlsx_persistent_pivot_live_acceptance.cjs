'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const pivotAgent=require('./xlsx-agent-pivot-task.cjs')
const pivotPersistent=require('./xlsx-persistent-pivot.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-persistent-pivot-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function liveApi(api){return {...api,pivotObserved:async(spec,apply)=>{const r=await pivotPersistent.runCommand(api.session,spec,apply);if(apply&&r?.ok&&!r.noOp)api.session.markWrite();return r}}}
function createTask(sheet,name,destination){return {operations:[{intent:'create_pivot',name,sourceSheet:sheet,sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:10},{items:['East','B'],expected:30},{items:['West','B'],expected:20}],...(destination?{pivotSheet:destination,destinationRange:'A1',repairIncompletePivot:true}:{})}]}}
function deleteTask(name,pivotSheet){return {operations:[{intent:'delete_pivot_sheet',name,pivotSheet}]}}
async function firstInvocation(api,sheet,name){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'pivot-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const values=[['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]
 const written=await api.writeRangeVerified({sheet,range:'A1:C5',values});if(!written.ok)return {ok:false,outcome:'pivot-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 return pivotAgent.executePivotTask({task:createTask(sheet,name),api:liveApi(api)})
}

async function existingSheetInvocation(api,sheet,destination,name){
 const created=await api.createSheetVerified(sheet);if(!created.ok)return {ok:false,outcome:'pivot-fixture-create-failed',authority:created.authority||'LIVE_READ',created}
 const dest=await api.createSheetVerified(destination);if(!dest.ok)return {ok:false,outcome:'pivot-destination-create-failed',authority:dest.authority||'LIVE_READ',dest}
 const values=[['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]
 const written=await api.writeRangeVerified({sheet,range:'A1:C5',values});if(!written.ok)return {ok:false,outcome:'pivot-fixture-write-failed',authority:written.authority||'LIVE_READ',written}
 return pivotAgent.executePivotTask({task:createTask(sheet,name,destination),api:liveApi(api)})
}
function exactCreate(r,name){const s=r?.wholeTaskVerification?.state;return s?.present===true&&s?.name===name&&s?.source==='A1:C5'&&s?.rowFields===1&&s?.columnFields===1&&s?.dataFields===1&&s?.styleName==='PivotStyleMedium2'&&s?.assertions?.every(x=>x.ok&&x.value===x.expected)}
function exactDelete(r,name){const s=r?.wholeTaskVerification?.state;return s?.pivot?.present===false&&s?.pivot?.name===name&&s?.pivotSheetPresent===false}
function d(r){return {ok:r.ok,outcome:r.outcome,authority:r.authority,noOp:r.noOp,oneEditorSession:r.persistentSession?.oneEditorSession,writes:r.persistentSession?.writes,barrier:r.persistentSession?.persistenceBarrier?.ok,wholeTaskVerification:r.wholeTaskVerification,diagnostic:r.ok?undefined:{pre:r.pre,applied:r.applied,final:r.final}}}
;(async()=>{
 const suffix=String(Date.now()).slice(-7),sheet=`EURO_PIV_${suffix}`,name=`EURO_P_${suffix}`
 const first=await persistent.withPersistentXlsxSession(options,api=>firstInvocation(api,sheet,name));console.log('PIVOT CREATE TASK',JSON.stringify(d(first),null,2))
 assert.equal(first.ok,true);assert.equal(first.authority,'LIVE_VERIFY');assert.equal(first.noOp,false);assert.equal(exactCreate(first,name),true);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
 const pivotSheet=first.wholeTaskVerification.state.parentSheet;assert.equal(typeof pivotSheet,'string')
 const retry=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:createTask(sheet,name),api:liveApi(api)}));console.log('PIVOT CREATE RETRY',JSON.stringify(d(retry),null,2))
 assert.equal(retry.ok,true);assert.equal(retry.authority,'LIVE_VERIFY');assert.equal(retry.noOp,true);assert.equal(exactCreate(retry,name),true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
 const deleted=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:deleteTask(name,pivotSheet),api:liveApi(api)}));console.log('PIVOT DELETE TASK',JSON.stringify(d(deleted),null,2))
 assert.equal(deleted.ok,true);assert.equal(deleted.authority,'LIVE_VERIFY');assert.equal(deleted.noOp,false);assert.equal(exactDelete(deleted,name),true);assert.equal(deleted.persistentSession?.writes,1);assert.equal(deleted.persistentSession?.persistenceBarrier?.ok,true)
 const deleteRetry=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:deleteTask(name,pivotSheet),api:liveApi(api)}));console.log('PIVOT DELETE RETRY',JSON.stringify(d(deleteRetry),null,2))
 assert.equal(deleteRetry.ok,true);assert.equal(deleteRetry.authority,'LIVE_VERIFY');assert.equal(deleteRetry.noOp,true);assert.equal(exactDelete(deleteRetry,name),true);assert.equal(deleteRetry.persistentSession?.writes,0);assert.equal(deleteRetry.persistentSession?.persistenceBarrier,null)
 const exSheet=`EURO_PIVX_${suffix}`,exDest=`EURO_PVDX_${suffix}`,exName=`EURO_PX_${suffix}`
 const existing=await persistent.withPersistentXlsxSession(options,api=>existingSheetInvocation(api,exSheet,exDest,exName));console.log('PIVOT EXISTING-SHEET CREATE TASK',JSON.stringify(d(existing),null,2))
 assert.equal(existing.ok,true);assert.equal(existing.authority,'LIVE_VERIFY');assert.equal(existing.noOp,false);assert.equal(exactCreate(existing,exName),true);assert.equal(existing.wholeTaskVerification.state.parentSheet,exDest);assert.equal(existing.persistentSession?.oneEditorSession,true);assert.equal(existing.persistentSession?.writes,4);assert.equal(existing.persistentSession?.persistenceBarrier?.ok,true)
 const existingRetry=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:createTask(exSheet,exName,exDest),api:liveApi(api)}));console.log('PIVOT EXISTING-SHEET RETRY',JSON.stringify(d(existingRetry),null,2))
 assert.equal(existingRetry.ok,true);assert.equal(existingRetry.authority,'LIVE_VERIFY');assert.equal(existingRetry.noOp,true);assert.equal(exactCreate(existingRetry,exName),true);assert.equal(existingRetry.persistentSession?.writes,0);assert.equal(existingRetry.persistentSession?.persistenceBarrier,null)
 const existingDeleted=await persistent.withPersistentXlsxSession(options,api=>pivotAgent.executePivotTask({task:deleteTask(exName,exDest),api:liveApi(api)}));assert.equal(existingDeleted.ok,true);assert.equal(existingDeleted.persistentSession?.persistenceBarrier?.ok,true)
  // Active-sheet negative control: deliberately leave an unrelated sheet active
 // immediately before existing-sheet pivot creation. The observer must bind
 // insertion to the requested destination and verify that activation.
 const acSheet=`EURO_PAC_${suffix}`,acDest=`EURO_PAD_${suffix}`,acOther=`EURO_PAO_${suffix}`,acName=`EURO_PAN_${suffix}`
 const activeContext=await persistent.withPersistentXlsxSession(options,async api=>{
  for(const s of [acSheet,acDest,acOther]){const cr=await api.createSheetVerified(s);if(!cr.ok)return cr}
  const values=[['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]
  const w=await api.writeRangeVerified({sheet:acSheet,range:'A1:C5',values});if(!w.ok)return w
  const activeBody=`try{var s=Api.GetSheet(${JSON.stringify(acOther)});if(!s||typeof s.SetActive!=='function')return {ok:false};s.SetActive();var a=Api.GetActiveSheet();return {ok:true,name:a&&a.GetName?a.GetName():null}}catch(err){return {ok:false,error:String(err)}}`
  const activated=await api.session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:api.session.apiWhere,body:activeBody})
  if(!activated?.ok||activated.name!==acOther)return {ok:false,outcome:'pivot-active-context-fixture-failed',authority:'LIVE_READ',activated}
  return pivotAgent.executePivotTask({task:createTask(acSheet,acName,acDest),api:liveApi(api)})
 });console.log('PIVOT ACTIVE-CONTEXT CREATE TASK',JSON.stringify(d(activeContext),null,2))
 assert.equal(activeContext.ok,true);assert.equal(activeContext.authority,'LIVE_VERIFY');assert.equal(activeContext.wholeTaskVerification.state.parentSheet,acDest);assert.equal(activeContext.persistentSession?.persistenceBarrier?.ok,true)

 // Exact exam-order interaction probe: sort + active filter + validation +
 // defined-name source identity + conditional formatting all precede pivot.
 // This targets the remaining semantic difference from the 109-op batch.
 const ex2Sheet=`EURO_PFI_${suffix}`,ex2Dest=`EURO_PDI_${suffix}`,ex2Name=`EURO_PII_${suffix}`,ex2Marker=`EURO_PS_${suffix}`
 const interaction=await persistent.withPersistentXlsxSession(options,async api=>{
  const a=await api.createSheetVerified(ex2Sheet);if(!a.ok)return a
  const b=await api.createSheetVerified(ex2Dest);if(!b.ok)return b
  const values=[['Month','Region','Country','Product','Units']]
  const regions=['Northern Europe','Western Europe','Southern Europe','Central Europe'],products=['Atlas Cloud','Orion Analytics','Vertex Security']
  for(let ri=0;ri<4;ri++)for(let pi=0;pi<3;pi++)for(let mi=0;mi<12;mi++)values.push(['M'+(mi+1),regions[ri],['Sweden','Germany','Italy','Poland'][ri],products[pi],10+ri*7+pi*3+mi])
  const w=await api.writeRangeVerified({sheet:ex2Sheet,range:'A1:E145',values});if(!w.ok)return w
  const sort=require('./xlsx-persistent-sort.cjs'),filter=require('./xlsx-persistent-filter.cjs'),dn=require('./xlsx-persistent-defined-name.cjs'),cf=require('./xlsx-persistent-conditional-format.cjs')
  async function family(mod,spec){const r=await mod.runCommand(api.session,spec,true);if(r?.ok&&!r.noOp)api.session.markWrite();return r}
  let r=await family(sort,{sheet:ex2Sheet,range:'A1:E145',keyRange:'B1:B145',order:'asc',hasHeaders:true});if(!r.ok)return r
  r=await family(filter,{intent:'filter_range',sheet:ex2Sheet,range:'A1:E145',field:2,criteria1:'Northern Europe',operator:'xlOr'});if(!r.ok)return r
  r=await family(dn,{intent:'set_defined_name',name:ex2Marker,refersTo:`=${ex2Sheet}!$A$1:$E$145`});if(!r.ok)return r
  r=await family(cf,{type:'cf.add',sheet:ex2Sheet,range:'E2:E145',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'0',priority:1,fillColor:[214,239,220]}});if(!r.ok)return r
  const assertions=[];for(let ri=0;ri<4;ri++)for(let pi=0;pi<3;pi++){let sum=0;for(let mi=0;mi<12;mi++)sum+=10+ri*7+pi*3+mi;assertions.push({items:[regions[ri],products[pi]],expected:sum})}
  return pivotAgent.executePivotTask({task:{operations:[{intent:'create_pivot',name:ex2Name,sourceSheet:ex2Sheet,sourceRange:'A1:E145',sourceIdentityName:ex2Marker,rowField:'Region',columnField:'Product',dataField:'Units',styleName:'PivotStyleMedium2',pivotSheet:ex2Dest,destinationRange:'A1',repairIncompletePivot:true,assertions}]},api:liveApi(api)})
 });console.log('PIVOT INTERACTION CREATE TASK',JSON.stringify(d(interaction),null,2))
 assert.equal(interaction.ok,true);assert.equal(interaction.authority,'LIVE_VERIFY');assert.equal(interaction.wholeTaskVerification.state.parentSheet,ex2Dest);assert.equal(interaction.persistentSession?.persistenceBarrier?.ok,true)

 // Large-prefix reproduction: the exam failure happens only after substantial
 // same-session worksheet work. Exercise the same pivot contract after a
 // deterministic 144-row source plus unrelated verified writes in this session.
 const pxSheet=`EURO_PFX_${suffix}`,pxDest=`EURO_PFD_${suffix}`,pxName=`EURO_PFP_${suffix}`
 const prefix=await persistent.withPersistentXlsxSession(options,async api=>{
  const a=await api.createSheetVerified(pxSheet);if(!a.ok)return a
  const b=await api.createSheetVerified(pxDest);if(!b.ok)return b
  const values=[['Month','Region','Country','Product','Units']]
  const regions=['Northern Europe','Western Europe','Southern Europe','Central Europe'],products=['Atlas Cloud','Orion Analytics','Vertex Security']
  for(let ri=0;ri<4;ri++)for(let pi=0;pi<3;pi++)for(let mi=0;mi<12;mi++)values.push(['M'+(mi+1),regions[ri],['Sweden','Germany','Italy','Poland'][ri],products[pi],10+ri*7+pi*3+mi])
  const w=await api.writeRangeVerified({sheet:pxSheet,range:'A1:E145',values});if(!w.ok)return w
  // Generate substantial same-session activity without changing pivot source semantics.
  for(let i=0;i<20;i++){const rr=await api.readRange({sheet:pxSheet,range:'A1:E145'});if(!rr.ok)return rr}
  const assertions=[];for(let ri=0;ri<4;ri++)for(let pi=0;pi<3;pi++){let sum=0;for(let mi=0;mi<12;mi++)sum+=10+ri*7+pi*3+mi;assertions.push({items:[regions[ri],products[pi]],expected:sum})}
  return pivotAgent.executePivotTask({task:{operations:[{intent:'create_pivot',name:pxName,sourceSheet:pxSheet,sourceRange:'A1:E145',rowField:'Region',columnField:'Product',dataField:'Units',styleName:'PivotStyleMedium2',pivotSheet:pxDest,destinationRange:'A1',repairIncompletePivot:true,assertions}]},api:liveApi(api)})
 });console.log('PIVOT LARGE-PREFIX CREATE TASK',JSON.stringify(d(prefix),null,2))
 assert.equal(prefix.ok,true);assert.equal(prefix.authority,'LIVE_VERIFY');assert.equal(prefix.wholeTaskVerification.state.parentSheet,pxDest);assert.equal(prefix.persistentSession?.persistenceBarrier?.ok,true)
  console.log('XLSX PERSISTENT PIVOT LIVE ACCEPTANCE: PASS')
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
