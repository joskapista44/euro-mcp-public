'use strict'
const assert=require('assert')
const {withPersistentXlsxSession}=require('./xlsx-persistent-session.cjs')
const FILE_ID=Number(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
process.env.EURO_PLAYWRIGHT_PATH=process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-vertical-alignment-runtime-probe');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const credentials={url:process.env.EURO_NEXTCLOUD_URL||'https://mt-server.eu',user:process.env.EURO_NEXTCLOUD_USER||'elliot',pass:secret('Elliot_nc_pass')}
function probeCommand(sheetName){
  function has(o,n){return !!o&&typeof o[n]==='function'}
  try{
    var sheet=Api.GetSheet(sheetName),range=sheet&&sheet.GetRange('A1:B2')
    if(!range||!has(range,'SetAlignVertical')||!range.range||!has(range.range,'getAlign'))return {ok:false,outcome:'vertical-alignment-probe-unavailable'}
    var values=['bottom','center','top','distributed','justify'],results=[]
    for(var i=0;i<values.length;i++){
      range.SetAlignVertical(values[i])
      var a=range.range.getAlign()
      results.push({requested:values[i],actual:a?a.ver:null,align:a||null})
    }
    return {ok:true,outcome:'vertical-alignment-runtime-measured',source:'live-coedit-editor',results:results}
  }catch(err){return {ok:false,outcome:'vertical-alignment-probe-error',error:String(err&&err.message||err)}}
}
async function runCommand(session,sheet){
  const body=`return (${probeCommand.toString()}).apply(null, ${JSON.stringify([sheet])});`
  return session.frame.evaluate(({where,body})=>new Promise(resolve=>{
    const e=where==='window.editor'?window.editor:(window.Asc||{}).editor
    let done=false
    const finish=v=>{if(!done){done=true;resolve(v)}}
    const timer=setTimeout(()=>finish({ok:false,outcome:'callback-timeout'}),5000)
    try{e.callCommand(new Function(body),false,v=>{clearTimeout(timer);finish(v)})}
    catch(err){clearTimeout(timer);finish({ok:false,outcome:'callcommand-error',error:String(err&&err.message||err)})}
  }),{where:session.apiWhere,body})
}
;(async()=>{
  const sheet=`EURO VAL ${String(Date.now()).slice(-7)}`
  const result=await withPersistentXlsxSession({url:credentials.url,user:credentials.user,pass:credentials.pass,fileId:FILE_ID,timeoutMs:30000,pollMs:50},async api=>{
    const created=await api.createSheetVerified(sheet)
    assert.equal(created.ok,true)
    const measured=await runCommand(api.session,sheet)
    const deleted=await api.deleteSheetVerified(sheet)
    return {ok:measured.ok&&deleted.ok,outcome:measured.outcome,authority:measured.ok&&deleted.ok?'LIVE_VERIFY':'LIVE_READ',noOp:false,created,measured,deleted}
  })
  console.log('XLSX VERTICAL ALIGNMENT RUNTIME PROBE',JSON.stringify({ok:result.ok,outcome:result.outcome,authority:result.authority,writes:result.persistentSession?.writes,barrier:result.persistentSession?.persistenceBarrier?.ok,results:result.measured?.results||null},null,2))
  assert.equal(result.ok,true)
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
