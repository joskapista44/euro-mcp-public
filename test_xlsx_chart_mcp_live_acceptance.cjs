'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-chart-mcp-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const env={...process.env,EURO_AGENT_ID:process.env.EURO_AGENT_ID||'elliot',EURO_COEDIT_AGENTS:process.env.EURO_COEDIT_AGENTS||'elliot',EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:secret('Elliot_nc_pass'),EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'}
function payload(r){const p=JSON.parse(r.content[0].text),failed=p.steps?.find(s=>s.result?.ok!==true)?.result;console.log(JSON.stringify({ok:p.ok,outcome:p.outcome,noOp:p.noOp,callerId:p.callerId,session:p.persistentSession||p.session,wholeTaskVerification:p.wholeTaskVerification,diagnostic:p.ok?undefined:{stepOutcome:failed?.outcome,pre:failed?.pre?.state,fresh:failed?.fresh?.state,applied:failed?.applied,final:failed?.final}},null,2));return p}
;(async()=>{
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env,stderr:'inherit'}),client=new Client({name:'xlsx-chart-live',version:'1'})
 await client.connect(transport)
 try{
  const suffix=String(Date.now()).slice(-7),sheet=`EURO_MC_${suffix}`,name=`EURO_C_${suffix}`,renamed=`EURO_R_${suffix}`
  const chartState={sheet,chartType:'bar',title:'Revenue',width:3600000,height:2160000,expectedSeriesCount:1,presentation:{legendPosition:'bottom',horizontalAxisTitle:'Month',verticalAxisTitle:'Revenue',dataLabels:{showSeriesName:false,showCategoryName:false,showValue:true,showPercent:false},style:2},chartPosition:{fromCol:4,colOffset:0,fromRow:1,rowOffset:0},series:[{index:0,name:'Verified Revenue',valuesRange:`${sheet}!$B$2:$B$3`,categoryRange:`${sheet}!$A$2:$A$3`}]}
  const setOps=[{intent:'create_sheet',name:sheet},{intent:'write_range',sheet,range:'A1:B3',values:[['Month','Revenue'],['Jan',10],['Feb',20]]},{intent:'set_chart',name,range:'A1:B3',...chartState}]
  console.log('CHART MCP APPLY');const first=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:setOps}},undefined,{timeout:90000}));assert.equal(first.ok,true);assert.equal(first.noOp,false);assert.equal(first.persistentSession?.oneEditorSession,true);assert.equal(first.persistentSession?.writes,3);assert.equal(first.persistentSession?.persistenceBarrier?.ok,true)
  console.log('CHART MCP RETRY');const retry=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:setOps}},undefined,{timeout:90000}));assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null)
  const renameOps=[{intent:'rename_chart',name,newName:renamed,...chartState}]
  console.log('CHART MCP RENAME');const rename=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:renameOps}},undefined,{timeout:90000}));assert.equal(rename.ok,true);assert.equal(rename.noOp,false);assert.equal(rename.persistentSession?.writes,1);assert.equal(rename.persistentSession?.persistenceBarrier?.ok,true)
  console.log('CHART MCP RENAME RETRY');const renameRetry=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:renameOps}},undefined,{timeout:90000}));assert.equal(renameRetry.ok,true);assert.equal(renameRetry.noOp,true);assert.equal(renameRetry.persistentSession?.writes,0);assert.equal(renameRetry.persistentSession?.persistenceBarrier,null)
  const delOps=[{intent:'delete_chart',sheet,name:renamed}]
  console.log('CHART MCP DELETE');const deleted=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:delOps}},undefined,{timeout:90000}));assert.equal(deleted.ok,true);assert.equal(deleted.noOp,false);assert.equal(deleted.persistentSession?.writes,1);assert.equal(deleted.persistentSession?.persistenceBarrier?.ok,true)
  console.log('CHART MCP DELETE RETRY');const deleteRetry=payload(await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:delOps}},undefined,{timeout:90000}));assert.equal(deleteRetry.ok,true);assert.equal(deleteRetry.noOp,true);assert.equal(deleteRetry.persistentSession?.writes,0);assert.equal(deleteRetry.persistentSession?.persistenceBarrier,null)
  console.log('XLSX CHART MCP LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
