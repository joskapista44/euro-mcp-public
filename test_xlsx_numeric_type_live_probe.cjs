'use strict'

// Diagnostic only: no worksheet changes, no inferred numeric storage type.
function probeCommand(sheetName) {
  function attempt(fn) {
    try { const value=fn(); return {ok:true,value:value===undefined?null:value,jsType:typeof value} }
    catch (_) { return {ok:false,error:'public-call-failed'} }
  }
  const funcs=Api.WorksheetFunction
  function measure(argument) {
    const out={}
    for(const name of ['TYPE','ISNUMBER','ISTEXT','COUNT','SUM']) {
      out[name]=funcs&&typeof funcs[name]==='function'
        ? attempt(()=>funcs[name](argument)) : {ok:false,error:'unavailable'}
    }
    return out
  }
  const controls=[120,'120',true].map(value=>({input:value,jsType:typeof value,results:measure(value)}))
  const typeControlsPass=controls.every((c,i)=>c.results.TYPE.ok&&c.results.TYPE.value===[1,2,4][i])
  const sheet=Api.GetSheet(sheetName)
  if(!sheet)return {ok:false,outcome:'probe-sheet-not-found',controls,typeControlsPass}
  const cells=['D1','E1','H1','D2','E2','H2'].map(address=>{
    const range=sheet.GetRange(address),getters={}
    for(const name of ['GetValue','GetValue2','GetText','GetFormula','GetNumberFormat']) {
      getters[name]=typeof range[name]==='function'?attempt(()=>range[name]()):{ok:false,error:'unavailable'}
    }
    return {address,getters,reference:measure(range),valueArgument:measure(getters.GetValue.value)}
  })
  return {ok:true,outcome:'diagnostic-collected-not-acceptance',source:'live-coedit-editor',sheet:sheetName,controls,typeControlsPass,cells}
}

async function run() {
  const caller=require('./coedit.cjs').detectCallerId()
  if(!caller.ok||caller.id!=='elliot')throw Error('allowlisted elliot required')
  const {getSecret}=require('/home/user/marveen/dist/web/vault.js')
  const pass=getSecret('Elliot_nc_pass','xlsx-numeric-type-readonly-probe')
  if(!pass)throw Error('vault credential missing')
  process.env.EURO_PLAYWRIGHT_PATH ||= '/home/user/marveen/node_modules/playwright'
  const {openMinimalXlsxSession}=require('./xlsx-minimal-editor-session.cjs')
  const session=await openMinimalXlsxSession({url:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',user:'elliot',pass,fileId:process.env.EURO_XLSX_FILE_ID||'1236770',timeoutMs:30000})
  let result
  try {
    const body=`return (${probeCommand.toString()})(${JSON.stringify(process.env.EURO_XLSX_TEST_SHEET||'MCP_Data_768SVZ')});`
    result=await session.frame.evaluate(({u,body})=>new Promise(resolve=>{
      const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor
      let settled=false
      const finish=value=>{if(!settled){settled=true;clearTimeout(timer);resolve(value)}}
      const timer=setTimeout(()=>finish({ok:false,outcome:'probe-callback-timeout'}),10000)
      try{editor.callCommand(new Function(body),false,value=>finish(value||{ok:false,outcome:'empty-callback'}))}
      catch(_){finish({ok:false,outcome:'probe-dispatch-failed'})}
    }),{u:session.apiWhere,body})
  } finally { await session.close() }
  console.error('NUMERIC TYPE PROBE',JSON.stringify({ok:result?.ok,outcome:result?.outcome,typeControlsPass:result?.typeControlsPass,oneEditorSession:true,writes:session.writes,closed:session.closed}))
  for(const control of result?.controls||[])console.error('CONTROL',JSON.stringify(control))
  for(const cell of result?.cells||[])console.error('CELL',JSON.stringify(cell))
  if(!result?.ok)process.exitCode=1
}

module.exports={probeCommand}
if(require.main===module)run().catch(()=>{console.error('Numeric probe failed; no acceptance claimed.');process.exitCode=1})
