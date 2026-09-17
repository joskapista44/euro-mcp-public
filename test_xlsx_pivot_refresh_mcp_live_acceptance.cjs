'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-pivot-refresh-mcp-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const env={...process.env,EURO_AGENT_ID:process.env.EURO_AGENT_ID||'elliot',EURO_COEDIT_AGENTS:process.env.EURO_COEDIT_AGENTS||'elliot',EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:secret('Elliot_nc_pass'),EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'}
function payload(label,r){const p=JSON.parse(r.content[0].text),failed=p.steps?.find(s=>s.result?.ok!==true)?.result;console.log(label+' '+JSON.stringify({ok:p.ok,outcome:p.outcome,noOp:p.noOp,session:p.persistentSession||p.session,wholeTaskVerification:p.wholeTaskVerification,diagnostic:p.ok?undefined:{stepOutcome:failed?.outcome,pre:failed?.pre?.state,fresh:failed?.fresh?.state,applied:failed?.applied,final:failed?.final}},null,2));return p}
function pivotSpec(intent,name,sheet,price){return {intent,name,sourceSheet:sheet,sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:price},{items:['East','B'],expected:30},{items:['West','B'],expected:20}]}}
;(async()=>{
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env,stderr:'inherit'}),client=new Client({name:'xlsx-pivot-refresh-live',version:'1'})
 await client.connect(transport)
 try{
  const suffix=String(Date.now()).slice(-7),sheet=`EURO_PR_${suffix}`,name=`EURO_P_${suffix}`
  const setupOps=[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:C5',values:[['Region','Style','Price'],['East','A',10],['West','B',20],['East','B',30],['West','A',40]]},pivotSpec('create_pivot',name,sheet,10)]
  const setup=payload('PIVOT REFRESH SETUP',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:setupOps}},undefined,{timeout:90000}));assert.equal(setup.ok,true);assert.equal(setup.persistentSession?.oneEditorSession,true)
  const pivotSheet=setup.steps?.find(s=>s.intent==='create_pivot')?.result?.wholeTaskVerification?.state?.parentSheet;assert.equal(typeof pivotSheet,'string')
  const refreshOps=[{intent:'write_range',sheet,range:'C2',values:[[15]]},pivotSpec('refresh_pivot',name,sheet,15)]
  const refreshed=payload('PIVOT REFRESH APPLY',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:refreshOps}},undefined,{timeout:90000}));assert.equal(refreshed.ok,true);assert.equal(refreshed.noOp,false);assert.equal(refreshed.persistentSession?.oneEditorSession,true);assert.equal(refreshed.persistentSession?.writes,2);assert.equal(refreshed.persistentSession?.persistenceBarrier?.ok,true)
  const retry=payload('PIVOT REFRESH RETRY',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:refreshOps}},undefined,{timeout:90000}));assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
  const deleted=payload('PIVOT REFRESH CLEANUP',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:[{intent:'delete_pivot_sheet',name,pivotSheet}]}},undefined,{timeout:90000}));assert.equal(deleted.ok,true);assert.equal(deleted.persistentSession?.writes,1)
  console.log('XLSX PIVOT REFRESH MCP LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
