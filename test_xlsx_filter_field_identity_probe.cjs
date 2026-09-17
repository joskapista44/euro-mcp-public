'use strict'
const assert=require('assert')
const persistent=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-filter-field-identity-probe');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
const options={url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50}
function filterProbeCommand(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function val(f,m,p){try{return has(f,m)?f[m]():f[p]}catch(e){return {error:String(e&&e.message||e)}}}
 try{
  var sh=Api.GetSheet(spec.sheet),range=sh&&sh.GetRange(spec.range)
  if(!range||!has(range,'SetAutoFilter'))return {ok:false,outcome:'filter-api-unavailable'}
  range.SetAutoFilter(spec.field,spec.criteria1,spec.operator)
  var af=sh.GetAutoFilter(),rg=af&&has(af,'GetRange')?af.GetRange():null,filters=af&&has(af,'GetFilters')?af.GetFilters():null
  if(!Array.isArray(filters))return {ok:false,outcome:'filters-not-array',range:rg&&has(rg,'GetAddress')?rg.GetAddress():null}
  var out=[]
  for(var i=0;i<filters.length;i++){
   var f=filters[i],own={},keys=[]
   try{keys=Object.keys(f||{})}catch(_){}
   for(var k=0;k<keys.length;k++){var name=keys[k],v;try{v=f[name]}catch(_){continue}if(v===null||['string','number','boolean'].includes(typeof v))own[name]=v}
   var methods=[];try{var p=Object.getPrototypeOf(f);if(p)methods=Object.getOwnPropertyNames(p).filter(function(n){return typeof f[n]==='function'})}catch(_){}
   out.push({arrayIndex:i,on:val(f,'GetOn','On'),operator:val(f,'GetOperator','Operator'),criteria1:val(f,'GetCriteria1','Criteria1'),criteria2:val(f,'GetCriteria2','Criteria2'),ownKeys:keys,primitiveOwn:own,prototypeMethods:methods})
  }
  return {ok:true,outcome:'filter-field-identity-probed',source:'live-coedit-editor',requested:spec,range:rg&&has(rg,'GetAddress')?rg.GetAddress():null,filterMode:af&&has(af,'GetFilterMode')?af.GetFilterMode():null,filterCount:filters.length,filters:out}
 }catch(e){return {ok:false,outcome:'filter-probe-error',error:String(e&&e.stack||e)}}
}
async function runProbe(session,spec){
 const body=`return (${filterProbeCommand.toString()})(${JSON.stringify(spec)});`
 return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout'}),5000);try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v)})}catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',error:String(err&&err.message||err)})}}),{where:session.apiWhere,body})
}
;(async()=>{
 const sheet=`EURO FLT PROBE ${String(Date.now()).slice(-6)}`
 const result=await persistent.withPersistentXlsxSession(options,async api=>{
  const created=await api.createSheetVerified(sheet);if(!created.ok)return created
  const written=await api.writeRangeVerified({sheet,range:'A1:C5',values:[['Name','Score','Group'],['Delta',40,'B'],['Alpha',10,'A'],['Charlie',30,'A'],['Bravo',20,'B']]});if(!written.ok)return written
  const observation=await runProbe(api.session,{sheet,range:'A1:C5',field:3,criteria1:'A',operator:'xlOr'})
  if(!observation.ok)return {ok:false,outcome:'filter-field-identity-probe-failed',authority:'LIVE_READ',observation}
  api.session.markWrite()
  return {ok:true,outcome:'filter-field-identity-probe-complete',authority:'LIVE_VERIFY',noOp:false,observation}
 })
 console.log('FILTER_FIELD_IDENTITY_PROBE='+JSON.stringify({ok:result.ok,outcome:result.outcome,observation:result.observation,persistentSession:result.persistentSession},null,2))
 assert.equal(result.ok,true);assert.equal(result.persistentSession?.oneEditorSession,true);assert.equal(result.persistentSession?.writes,3);assert.equal(result.persistentSession?.persistenceBarrier?.ok,true)
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
