'use strict'
const assert=require('assert/strict')
const path=require('path')
const {Client}=require('@modelcontextprotocol/sdk/client/index.js')
const {StdioClientTransport}=require('@modelcontextprotocol/sdk/client/stdio.js')
;(async()=>{
 const caller=require('./coedit.cjs').detectCallerId()
 if(!caller.ok)throw Error('MCP caller is not configured/allowlisted')
 if(caller.id!=='elliot')throw Error('This acceptance fixture requires elliot')
 const vault=require('/home/user/marveen/dist/web/vault.js')
 const pass=vault.getSecret('Elliot_nc_pass','xlsx-batch-mcp-live-acceptance')
 if(!pass)throw Error('vault secret missing')
 const client=new Client({name:'xlsx-mcp-live-acceptance',version:'1'})
 const transport=new StdioClientTransport({command:process.execPath,args:[path.join(__dirname,'euro-mcp-m44.cjs')],cwd:__dirname,env:{...process.env,EURO_COEDIT_NC_URL:process.env.EURO_COEDIT_NC_URL||'https://mt-server.eu',ELLIOT_NEXTCLOUD_USER:'elliot',ELLIOT_NEXTCLOUD_APP_PASSWORD:pass,EURO_PLAYWRIGHT_PATH:process.env.EURO_PLAYWRIGHT_PATH||'/home/user/marveen/node_modules/playwright'},stderr:'inherit'})
 try{
  await client.connect(transport)
  assert((await client.listTools()).tools.some(t=>t.name==='office_xlsx_batch'))
  const suffix=Date.now(),source='EURO_MCP_SRC_'+suffix,sheet='EURO_MCP_'+suffix,name='EURO_MN_'+suffix,pivot='EURO_MP_'+suffix,file_id=String(process.env.EURO_XLSX_PERSISTENT_FILE_ID||1231187)
  const args={file_id,operations:[
   {intent:'create_sheet',name:source},
   {intent:'write_range',sheet:source,range:'A1:E5',values:[['Region','Style','Price','Flag','Scratch'],['East','A',10,'Y','remove'],['West','B',20,'N',''],['East','B',30,'Y',''],['West','A',40,'N','']]},
   {intent:'rename_sheet',sheet:source,name:sheet},
   {intent:'move_sheet',sheet,referenceSheet:'Sheet1',position:'before'},
   {intent:'clear_range',sheet,range:'E2'},
   {intent:'format_range',sheet,range:'A1:E1',format:{bold:true,fillColor:[210,220,230]}},
   {intent:'add_conditional_format',sheet,range:'C2:C5',rule:{type:'xlCellValue',operator:'xlGreater',formula1:'10',fillColor:[255,0,0],priority:1}},
   {intent:'set_defined_name',name,refersTo:'='+sheet+'!$A$1:$C$5'},
   {intent:'create_pivot',name:pivot,sourceSheet:sheet,sourceRange:'A1:C5',rowField:'Region',columnField:'Style',dataField:'Price',styleName:'PivotStyleMedium2',assertions:[{items:['East','A'],expected:10},{items:['East','B'],expected:30},{items:['West','B'],expected:20}]}
  ]}
  for(let i=0;i<2;i++){
   const reply=await client.callTool({name:'office_xlsx_batch',arguments:args},undefined,{timeout:90000})
   const r=JSON.parse(reply.content.find(c=>c.type==='text').text)
   console.log(i?'MCP PERSISTED RETRY':'MCP APPLY',JSON.stringify(r.ok?{ok:r.ok,noOp:r.noOp,callerId:r.callerId,session:r.persistentSession,wholeTaskVerification:r.wholeTaskVerification}:r,null,2))
   assert.equal(reply.isError,false);assert.equal(r.ok,true);assert.equal(r.authority,'LIVE_VERIFY');assert.equal(r.noOp,i===1)
   assert.equal(r.persistentSession.oneEditorSession,true);assert.equal(r.persistentSession.writes,i?0:8)
   assert.equal(r.wholeTaskVerification.checks.length,9)
   assert.equal(r.wholeTaskVerification.readOnly,true)
   if(i)assert.equal(r.persistentSession.persistenceBarrier,null)
   else assert.equal(r.persistentSession.persistenceBarrier.ok,true)
  }
  console.log('XLSX BATCH MCP LIVE ACCEPTANCE: PASS')
 }finally{await client.close()}
})().catch(e=>{console.error(e.stack||e);process.exitCode=1})
