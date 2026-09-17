'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
const FILE_ID=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
function secret(id){const v=require('/home/user/marveen/dist/web/vault.js');const r=v.getSecret(id,'xlsx-final-power-user-mcp-live-acceptance');if(!r)throw new Error(`vault secret not found: ${id}`);return r}
const env={...process.env,EURO_AGENT_ID:process.env.EURO_AGENT_ID||'elliot',EURO_COEDIT_AGENTS:process.env.EURO_COEDIT_AGENTS||'elliot',EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:secret('Elliot_nc_pass'),EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'}
function payload(label,r){const p=JSON.parse(r.content[0].text),failed=p.steps?.find(s=>s.result?.ok!==true)?.result;console.log(label+' '+JSON.stringify({ok:p.ok,outcome:p.outcome,noOp:p.noOp,session:p.persistentSession||p.session,wholeTaskVerification:p.wholeTaskVerification,diagnostic:p.ok?undefined:{stepOutcome:failed?.outcome,pre:failed?.pre?.state,fresh:failed?.fresh?.state,applied:failed?.applied,final:failed?.final}},null,2));return p}
;(async()=>{
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env,stderr:'inherit'}),client=new Client({name:'xlsx-final-power-user-live',version:'1'})
 await client.connect(transport)
 try{
  const suffix=String(Date.now()).slice(-7),data=`EURO_FD_${suffix}`,chartSheet=`EURO_FC_${suffix}`,definedName=`EURO_FN_${suffix}`,pivotName=`EURO_FP_${suffix}`,chartName=`EURO_FCH_${suffix}`
  const setupOps=[
   {intent:'create_sheet',name:data},{intent:'write_range',sheet:data,range:'A1:E5',values:[['Region','Style','Price','Flag','Scratch'],['West','B',20,'N','remove'],['East','A',10,'Y',''],['West','A',40,'N',''],['East','B',30,'Y','']]},
   {intent:'create_sheet',name:chartSheet},{intent:'write_range',sheet:chartSheet,range:'A1:B3',values:[['Month','Revenue'],['Jan',10],['Feb',20]]}
  ]
  const setup=payload('FINAL POWER USER SETUP',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:setupOps}},undefined,{timeout:90000}));assert.equal(setup.ok,true);assert.equal(setup.persistentSession?.oneEditorSession,true);assert.equal(setup.persistentSession?.writes,4)
  const chartState={sheet:chartSheet,name:chartName,range:'A1:B3',chartType:'bar',title:'Revenue',width:3600000,height:2160000,expectedSeriesCount:1,presentation:{legendPosition:'bottom',horizontalAxisTitle:'Month',verticalAxisTitle:'Revenue',dataLabels:{showSeriesName:false,showCategoryName:false,showValue:true,showPercent:false},style:2},chartPosition:{fromCol:4,colOffset:0,fromRow:1,rowOffset:0},series:[{index:0,name:'Verified Revenue',valuesRange:`${chartSheet}!$B$2:$B$3`,categoryRange:`${chartSheet}!$A$2:$A$3`}]}
  const operations=[
   {intent:'move_sheet',sheet:data,referenceSheet:'Sheet1',position:'before'},
   {intent:'clear_range',sheet:data,range:'E2'},
   {intent:'format_range',sheet:data,range:'A1:E1',format:{bold:true,fillColor:[210,220,230]}},
   {intent:'layout_range',sheet:data,range:'A1:A5',type:'column.width',width:24},
   {intent:'merge_range',sheet:data,range:'E6:F6'},
   {intent:'freeze_panes',sheet:data,mode:'at',range:'A2'},
   {intent:'sort_range',sheet:data,range:'A1:C5',keyRange:'A1:A5',order:'asc',hasHeaders:true},
   {intent:'filter_range',sheet:data,range:'A1:C5',field:2,criteria1:'A',operator:'xlOr'},
   {intent:'set_validation',sheet:data,range:'F2:F5',validationType:'xlValidateWholeNumber',alertStyle:'xlValidAlertStop',operator:'xlBetween',formula1:1,formula2:10},
   {intent:'set_defined_name',name:definedName,refersTo:`=${data}!$A$1:$C$5`},
   {intent:'add_conditional_format',sheet:data,range:'C2:C5',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:[255,0,0],priority:1}},
   {intent:'create_pivot',name:pivotName,sourceSheet:data,sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:10},{items:['East','B'],expected:30},{items:['West','B'],expected:20}]},
   {intent:'set_chart',...chartState}
  ]
  const applied=payload('FINAL POWER USER APPLY',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations}},undefined,{timeout:90000}));assert.equal(applied.ok,true);assert.equal(applied.noOp,false);assert.equal(applied.persistentSession?.oneEditorSession,true);assert.equal(applied.persistentSession?.writes,operations.length);assert.equal(applied.persistentSession?.persistenceBarrier?.ok,true);assert.equal(applied.wholeTaskVerification?.readOnly,true);assert.equal(applied.wholeTaskVerification?.checks?.length,operations.length)
  const pivotSheet=applied.steps?.find(s=>s.intent==='create_pivot')?.result?.wholeTaskVerification?.state?.parentSheet;assert.equal(typeof pivotSheet,'string')
  const retry=payload('FINAL POWER USER RETRY',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations}},undefined,{timeout:90000}));assert.equal(retry.ok,true);assert.equal(retry.noOp,true);assert.equal(retry.persistentSession?.oneEditorSession,true);assert.equal(retry.persistentSession?.writes,0);assert.equal(retry.persistentSession?.persistenceBarrier,null);assert.equal(retry.wholeTaskVerification?.checks?.length,operations.length)
  const cleanup=payload('FINAL POWER USER OBJECT CLEANUP',await client.callTool({name:'office_xlsx_batch',arguments:{file_id:FILE_ID,operations:[{intent:'delete_pivot_sheet',name:pivotName,pivotSheet},{intent:'delete_chart',sheet:chartSheet,name:chartName}]}},undefined,{timeout:90000}));assert.equal(cleanup.ok,true);assert.equal(cleanup.persistentSession?.writes,2)
  console.log('XLSX FINAL POWER USER MCP LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(e=>{console.error(e?.stack||e);process.exitCode=1})
